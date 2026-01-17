
import io
import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell
from odoo import http
from odoo.http import request, content_disposition
from collections import defaultdict
from odoo.exceptions import UserError, ValidationError, AccessError


class StockInOutReportController(http.Controller):

    @http.route('/web/binary/export_xlsx', type='http', auth='user')
    def export_xlsx(self, wizard_id=None, **kwargs):
        wizard = request.env['bk.stock.inout.wizard'].browse(int(wizard_id))
        if not wizard.exists():
            return request.not_found()

        output = io.BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        sheet = workbook.add_worksheet("Stock Report")

        # Formats
        title_fmt = workbook.add_format({'bold': True, 'font_size': 14, 'align': 'center'})
        header_fmt = workbook.add_format({
            'bold': True, 'bg_color': '#D9D9D9', 'border': 1,
            'align': 'center', 'valign': 'vcenter', 'text_wrap': True
        })
        number_fmt = workbook.add_format({'border': 1, 'font_size': 9, 'num_format': '#,##0.00', 'align': 'right'})
        alt_row_fmt = workbook.add_format({'border': 1, 'font_size': 9, 'num_format': '#,##0.00',
                                           'align': 'right', 'bg_color': '#F9F9F9'})
        total_fmt = workbook.add_format({'bold': True, 'bg_color': '#E0E0E0', 'border': 1,
                                         'num_format': '#,##0.00', 'align': 'right'})
        text_fmt = workbook.add_format({'border': 1, 'font_size': 9, 'align': 'left'})

        # 1. Report Title
        sheet.merge_range('A1:H1', "Stock Movement Balance Report", title_fmt)

        # 2. Filter Conditions
        sheet.write('A2', f"Period: {wizard.date_start} → {wizard.date_end}")
        sheet.write('A3', f"Locations: {', '.join([loc.display_name for loc in wizard.location_ids]) or 'All'}")
        sheet.write('A4', f"Products/Categories: "
                          f"{', '.join([p.display_name for p in wizard.product_ids]) if wizard.product_ids else ''}"
                          f"{', '.join([c.display_name for c in wizard.categ_ids]) if wizard.categ_ids else 'All'}")

        # 3. Data
        if wizard.report_type == 'detailed':
            headers, lines = self._compute_detailed_lines(wizard)
        elif wizard.report_type == "summary":  # summary
            headers, lines = self._compute_summary_lines(wizard)
        else:
            headers, lines = self._compute_summary_valuation_lines(wizard)

        # write headers
        for col, h in enumerate(headers):
            sheet.write(5, col, h, header_fmt)
            sheet.set_column(col, col, 14)  # set default width

        # write data rows
        row = 6
        for idx, line in enumerate(lines):
            row_fmt = alt_row_fmt if idx % 2 else number_fmt
            for col, val in enumerate(line):
                if isinstance(val, (int, float)):
                    sheet.write(row, col, val, row_fmt)
                else:
                    sheet.write(row, col, val, text_fmt)
            row += 1

        # optional totals row at end
        if lines and any(isinstance(v, (int, float)) for v in lines[0][3:]):  # detect numeric cols
            sheet.write(row, 0, "TOTAL", total_fmt)
            for col in range(3, len(headers)):
                col_letter = chr(65 + col)
                sheet.write_formula(row, col, f"SUM({col_letter}7:{col_letter}{row})", total_fmt)

        workbook.close()
        output.seek(0)

        file_name = "stock_report.xlsx"
        return request.make_response(
            output.read(),
            headers=[
                ('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'),
                ('Content-Disposition', content_disposition(file_name))
            ]
        )

    def _compute_lines(self, wizard):
        domain = [
            ("date", "<=", wizard.date_end),
        ]
        if wizard.state != 'all':
            domain.append(("state", "=", wizard.state))

        if wizard.product_ids:
            domain.append(("product_id", "in", wizard.product_ids.ids))
        elif wizard.categ_ids:
            domain.append(("product_id.categ_id", "in", wizard.categ_ids.ids))
        if wizard.location_ids:
            domain += ["|", ("location_id", "in", wizard.location_ids.ids),
                       ("location_dest_id", "in", wizard.location_ids.ids)]

        moves = request.env["stock.move"].sudo().search(domain, order="product_id, location_id, date")

        data = defaultdict(lambda: defaultdict(float))

        for move in moves:
            product = move.product_id
            categ = product.categ_id.display_name
            location = move.location_id if move.location_id.usage == "internal" else move.location_dest_id
            location_name = location.display_name if location else "-"

            key = (product.id, location.id)
            d = data[key]
            d["product"] = product.display_name
            d["category"] = categ
            d["location"] = location_name

            qty = move.product_uom_qty

            # Initial balance (before period)
            if move.date.date() < wizard.date_start:
                if move.location_dest_id.usage == "internal":
                    d["initial"] += qty
                if move.location_id.usage == "internal":
                    d["initial"] -= qty
                continue

            # Classification within period
            if move.location_id.usage == "supplier" and move.location_dest_id.usage == "internal":
                d["purchased"] += qty
            elif move.location_id.usage == "internal" and move.location_dest_id.usage == "supplier":
                d["return_sup"] += qty
            elif move.location_id.usage == "internal" and move.location_dest_id.usage == "customer":
                d["sold"] += qty
            elif move.location_id.usage == "customer" and move.location_dest_id.usage == "internal":
                d["return_cust"] += qty
            elif move.location_id.usage == "internal" and move.location_dest_id.usage in ("inventory", "production"):
                d["losses"] += qty
            elif move.location_id.usage in ("inventory", "production") and move.location_dest_id.usage == "internal":
                d["gains"] += qty

        # Compute ending and running balance
        lines = []
        for k, d in data.items():
            # Ending balance
            d["ending"] = (
                    d.get("initial", 0)
                    + d.get("purchased", 0)
                    + d.get("return_cust", 0)
                    + d.get("gains", 0)
                    - d.get("sold", 0)
                    - d.get("return_sup", 0)
                    - d.get("losses", 0)
            )
            # Running balance = initial + net movements per period
            d["running_balance"] = d["ending"]  # or keep same as ending

            # NEW: Net Incoming / Net Outgoing
            # NEW: Net Incoming / Net Outgoing
            # d["net_incoming"] = d.get("purchased", 0) + d.get("gains", 0) - d.get("losses", 0) - d.get("return_sup", 0)
            # d["net_outgoing"] = d.get("sold", 0) + d.get("losses", 0) - d.get("return_cust", 0)
            d["net_incoming"] = d.get("purchased", 0) + d.get("gains", 0) - d.get("return_sup", 0)
            d["net_outgoing"] = -1 * (d.get("sold", 0) + d.get("losses", 0) - d.get("return_cust", 0))

            lines.append(d)

        return lines

    def _compute_detailed_lines(self, wizard):
        """Detailed: per product + internal location with full breakdown."""
        # --- Build domain
        domain = [("date", "<=", wizard.date_end)]
        # state filter (ready→assigned)
        state_val = getattr(wizard, "state", "done")

        if state_val and state_val != "all":
            domain.append(("state", "=", "assigned" if state_val == "ready" else state_val))

        # product/category filter
        if wizard.product_ids:
            domain.append(("product_id", "in", wizard.product_ids.ids))
        elif wizard.categ_ids:
            domain.append(("product_id.categ_id", "in", wizard.categ_ids.ids))
        # location filter
        if wizard.location_ids:
            domain += ["|", ("location_id", "in", wizard.location_ids.ids),
                       ("location_dest_id", "in", wizard.location_ids.ids)]
        moves = request.env["stock.move"].sudo().search(domain, order="product_id, location_id, location_dest_id, date, id")
        # accumulators per (product, internal_location)
        agg = defaultdict(lambda: {
            "product": "",
            "status":"",
            "category": "",
            "location": "",
            "initial": 0.0,
            "purchased": 0.0,
            "return_sup": 0.0,
            "sold": 0.0,
            "return_cust": 0.0,
            "losses": 0.0,
            "gains": 0.0,
            "transfer_in":0.0,
            "transfer_out":0.0,
        })

        def pick_internal_loc(m):
            """Choose the internal location relevant for this move row."""
            src, dst = m.location_id, m.location_dest_id
            # If user filtered locations, prefer the one among them
            if wizard.location_ids:
                if src.usage == "internal" and src in wizard.location_ids:
                    return src
                if dst.usage == "internal" and dst in wizard.location_ids:
                    return dst
            # Otherwise pick whichever side is internal (if any)
            if src.usage == "internal":
                return src
            if dst.usage == "internal":
                return dst
            return None

        for mv in moves:
            loc = pick_internal_loc(mv)
            # Skip moves that don't touch an internal location at all
            if not loc:
                continue

            prod = mv.product_id
            key = (prod.id, loc.id)
            d = agg[key]
            d["product"] = prod.display_name
            d["status"] = prod.active
            d["category"] = prod.categ_id.display_name
            d["location"] = loc.display_name
            qty = mv.quantity #product_uom_qty

            # Initial balance (all moves before start date)
            if mv.date.date() < wizard.date_start:
                if mv.location_dest_id.usage == "internal":
                    d["initial"] += qty
                if mv.location_id.usage == "internal":
                    d["initial"] -= qty
                continue

            # Period classifications
            src_u = mv.location_id.usage
            dst_u = mv.location_dest_id.usage

            if src_u == "supplier" and dst_u == "internal":
                d["purchased"] += qty
            elif src_u == "internal" and dst_u == "supplier":
                d["return_sup"] += qty
            elif src_u == "internal" and dst_u == "customer":
                d["sold"] += qty
            elif src_u == "customer" and dst_u == "internal":
                d["return_cust"] += qty
            elif src_u == "internal" and dst_u in ("inventory", "production"):
                d["losses"] += qty
            elif src_u in ("inventory", "production") and dst_u == "internal":
                d["gains"] += qty
            elif src_u == "internal" and dst_u == "internal":
                # Count transfer out for source location
                key_src = (prod.id, mv.location_id.id)
                d_src = agg[key_src]
                d_src["product"] = prod.display_name
                d_src["category"] = prod.categ_id.display_name
                d_src["location"] = mv.location_id.display_name
                d_src["transfer_out"] += qty

                # Count transfer in for destination location
                key_dst = (prod.id, mv.location_dest_id.id)
                d_dst = agg[key_dst]
                d_dst["product"] = prod.display_name
                d_dst["category"] = prod.categ_id.display_name
                d_dst["location"] = mv.location_dest_id.display_name
                d_dst["transfer_in"] += qty

            elif src_u == "internal" and dst_u == "internal" and loc.id == mv.location_dest_id.id:
                d["transfer_in"] += qty
            elif src_u == "internal" and dst_u == "internal" and loc.id == mv.location_id.id:
                d["transfer_out"] += qty

        # finalize rows
        headers = [
            "Product", "Category", "Active", "Location","UoM",
            "Initial Balance",
            "Purchased", "Customer Returns", "Adjustments (Gain)",
            "Sold Qty", "Supplier Returns", "Adjustments (Loss)",
            "Transfer Issued", "Transfer Received",
            "Total Incoming", "Total Outgoing",
            "Ending Balance",
        ]
        lines = []
        running_total = 0.0  # overall running (sheet-level); keep if you want a per-line running figure
        for (pid, lid), d in agg.items():  # iterate with keys + values
            net_incoming = d["purchased"] + d["gains"] - d["return_sup"] + d["transfer_in"]
            net_outgoing = d["sold"] + d["losses"] - d["return_cust"] + d["transfer_out"]
            ending = d["initial"] + net_incoming - net_outgoing
            running_total += ending

            # fetch correct product (one record only)

            product = request.env["product.product"].browse(pid) if pid else False
            unit_cost = product.standard_price if product else 0.0
            uom = product.uom_id.name if product else ""
            # valuation = ending * unit_cost

            lines.append([
                d["product"], d["category"],
                d["status"],
                d["location"],
                uom,
                d["initial"],
                d["purchased"], d["return_cust"], d["gains"],
                d["sold"], d["return_sup"], d["losses"],
                d["transfer_out"],d["transfer_in"],
                net_incoming, net_outgoing,
                ending,
            ])

        return headers, lines

    def _compute_summary_lines(self, wizard):
        """Summary: per product (no location), with Initial, Net In, Net Out, Forecast."""
        # --- Build domain
        domain = [("date", "<=", wizard.date_end), ('product_id.active', '=', True )]
        # state filter (ready→assigned)
        state_val = getattr(wizard, "state", "done")
        if state_val and state_val != "all":
            domain.append(("state", "=", "assigned" if state_val == "ready" else state_val))
        # product/category filter
        if wizard.product_ids:
            domain.append(("product_id", "in", wizard.product_ids.ids))
        elif wizard.categ_ids:
            domain.append(("product_id.categ_id", "in", wizard.categ_ids.ids))

        # location filter (still applied for which moves are considered)
        if wizard.location_ids:
            domain += ["|", ("location_id", "in", wizard.location_ids.ids),
                            ("location_dest_id", "in", wizard.location_ids.ids)]

        moves = request.env["stock.move"].sudo().search(domain, order="product_id, date, id")

        # accumulators per product
        agg = defaultdict(lambda: {
            "product": "",
            "category": "",
            "initial": 0.0,
            "purchased": 0.0,
            "return_sup": 0.0,
            "sold": 0.0,
            "return_cust": 0.0,
            "losses": 0.0,
            "gains": 0.0,
        })

        for mv in moves:
            # consider only moves that touch any internal location so we don't count external-to-external noise
            if mv.location_id.usage != "internal" and mv.location_dest_id.usage != "internal":
                continue

            prod = mv.product_id
            d = agg[prod.id]
            d["product"] = prod.display_name
            d["category"] = prod.categ_id.display_name
            qty = mv.quantity #product_uom_qty
            src_u = mv.location_id.usage
            dst_u = mv.location_dest_id.usage

            # Initial balance (before start date)
            if mv.date.date() < wizard.date_start:
                if dst_u == "internal":
                    d["initial"] += qty
                if src_u == "internal":
                    d["initial"] -= qty
                continue

            # Period classifications (same as detailed)
            if src_u == "supplier" and dst_u == "internal":
                d["purchased"] += qty
            elif src_u == "internal" and dst_u == "supplier":
                d["return_sup"] += qty
            elif src_u == "internal" and dst_u == "customer":
                d["sold"] += qty
            elif src_u == "customer" and dst_u == "internal":
                d["return_cust"] += qty
            elif src_u == "internal" and dst_u in ("inventory", "production"):
                d["losses"] += qty
            elif src_u in ("inventory", "production") and dst_u == "internal":
                d["gains"] += qty

        headers = [
            "Product", "Category", "UoM",
            "Initial Balance",
            "Total Incoming", "Total Outgoing",
            "Forecast Qty",
        ]

        lines = []
        for pid, d in agg.items():
            net_incoming = d["purchased"] + d["gains"] - d["return_sup"]
            net_outgoing = d["sold"] + d["losses"] - d["return_cust"]
            forecast = d["initial"] + net_incoming - net_outgoing

            # fetch product to get UoM and cost
            product = request.env["product.product"].browse(pid) if pid else False
            uom = product.uom_id.name if product else ""
            # unit_cost = product.standard_price if product else 0.0
            # valuation = forecast * unit_cost

            lines.append([
                d["product"], d["category"],
                uom,  # 👈 add UoM here
                d["initial"],
                net_incoming, net_outgoing,
                forecast,
            ])
        return headers, lines

    def _compute_summary_valuation_lines(self, wizard):
        """Summary: per product (no location), with Initial, Net In, Net Out, Forecast."""
        domain = [("stock_move_id.date", "<=", wizard.date_end)]
        # state filter (ready→assigned)
        state_val = getattr(wizard, "state", "done")
        if state_val and state_val != "all":
            domain.append(("stock_move_id.state", "=", "assigned" if state_val == "ready" else state_val))

        # product/category filter
        if wizard.product_ids:
            domain.append(("product_id", "in", wizard.product_ids.ids))
        elif wizard.categ_ids:
            domain.append(("product_id.categ_id", "in", wizard.categ_ids.ids))

        # location filter
        if wizard.location_ids:
            domain += ["|", ("location_id", "in", wizard.location_ids.ids),
                       ("location_dest_id", "in", wizard.location_ids.ids)]

        # # include adjustments OR posted accounting moves
        # domain = domain
        # # [
        # #     "|",
        # #     ("stock_move_id.reference", "ilike", "Product Quantity"),
        # #     "&",
        # #     ("account_move_id", "!=", False),
        # #     ("account_move_id.state", "=", "posted"),
        # # ]

        moves = request.env["stock.valuation.layer"].sudo().search(domain)

        # discover inventory/COGS accounts
        acc_domain = ["|", ('is_cogs_account', '=', True), ('is_inventory_account', "=", True)]
        account_ids = request.env["account.account"].search(acc_domain)
        cogs_account = account_ids.filtered(lambda acc: acc.is_cogs_account)
        inventory_account = account_ids.filtered(lambda acc: acc.is_inventory_account)

        inventory_acc_ids = inventory_account.ids if inventory_account else []
        cogs_acc_ids = cogs_account.ids if cogs_account else []
        # ------------------------------------------------------
        # aggregator
        # ------------------------------------------------------
        from collections import defaultdict

        agg = defaultdict(lambda: {
            "product": "",
            "category": "",
            "status": "",
            "initial": 0.0,
            "purchased": 0.0,
            "return_sup": 0.0,
            "sold": 0.0,
            "return_cust": 0.0,
            "losses": 0.0,
            "gains": 0.0,
            "avg_sales_price": 0.0,
            "stock_valuation": 0.0,  # incoming value
            "cogs_value": 0.0,  # outgoing value

            "stock_valuation_account_id": "",
            "stock_cogs_account_id": "",
            "stock_income_account_id": "",
        })

        # ------------------------------------------------------
        # 1) QUANTITIES FROM SVL — KEEP YOUR LOGIC EXACTLY
        # ------------------------------------------------------
        processed_keys = set()

        for mv in moves:
            src_u = mv.stock_move_id.location_id.usage
            dst_u = mv.stock_move_id.location_dest_id.usage

            if src_u != "internal" and dst_u != "internal":
                continue

            prod = mv.product_id
            d = agg[prod.id]

            d["product"] = prod.display_name
            avg_price = self.get_avg_selling_price(mv.product_id, wizard.date_start, wizard.date_end, wizard.pos_config_ids.ids)
            d["avg_sales_price"] = avg_price if avg_price > 0 else mv.product_id.list_price
            d["status"] = prod.active
            d["category"] = prod.categ_id.display_name

            # account labels (unchanged)
            val_acc_id = prod.categ_id.property_stock_valuation_account_id
            cogs_acc_id = prod.categ_id.property_stock_account_output_categ_id
            inc_acc_id = prod.categ_id.property_account_income_categ_id

            d["stock_valuation_account_id"] = f"{val_acc_id.code} {val_acc_id.name}" if val_acc_id else ""
            d["stock_cogs_account_id"] = f"{cogs_acc_id.code} {cogs_acc_id.name}" if cogs_acc_id else ""
            d["stock_income_account_id"] = f"{inc_acc_id.code} {inc_acc_id.name}" if inc_acc_id else ""

            qty = mv.quantity or 0.0

            # initial qty pre-period
            if mv.stock_move_id.date.date() < wizard.date_start:
                if dst_u == "internal":
                    d["initial"] += qty
                if src_u == "internal":
                    d["initial"] -= qty
                continue

            # your movement classification untouched
            if src_u == "supplier" and dst_u == "internal":
                d["purchased"] += qty

            elif src_u == "internal" and dst_u == "supplier":
                d["return_sup"] += qty

            elif src_u == "internal" and dst_u == "customer":
                d["sold"] += qty

            elif src_u == "customer" and dst_u == "internal":
                d["return_cust"] += qty

            elif src_u == "internal" and dst_u in ("inventory", "production"):
                d["losses"] += qty

            elif src_u in ("inventory", "production") and dst_u == "internal":
                d["gains"] += qty

        # ------------------------------------------------------
        # 2) NEW CLEAN VALUATION: GET VALUES FROM ACCOUNT.MOVE.LINE
        # ------------------------------------------------------
        aml = request.env["account.move.line"].sudo().search([
            ("move_id.state", "=", "posted"),
            ("move_id.date", ">=", wizard.date_start),
            ("move_id.date", "<=", wizard.date_end),
            ("product_id", "!=", False),
            ("account_id", "in", inventory_acc_ids + cogs_acc_ids),
        ])

        valuation_map = defaultdict(lambda: {"incoming": 0.0, "outgoing": 0.0})

        for line in aml:
            pid = line.product_id.id
            bal = line.debit - line.credit

            if line.account_id.id in inventory_acc_ids:
                valuation_map[pid]["incoming"] += bal
            else:
                valuation_map[pid]["outgoing"] += bal

        # ------------------------------------------------------
        # 3) MERGE VALUE RESULTS INTO agg
        # ------------------------------------------------------
        for pid, d in agg.items():
            vals = valuation_map.get(pid, {})
            d["stock_valuation"] = vals.get("incoming", 0.0)
            d["cogs_value"] = abs(vals.get("outgoing", 0.0))

        # ------------------------------------------------------
        # 4) BUILD OUTPUT LINES
        # ------------------------------------------------------
        headers = [
            "Product", "Category", "UoM", "Active",
            "Initial Balance",
            "Total Incoming", "Total Outgoing",
            "Sold Qty",
            "Forecast Qty",
            "Valuation",
            "Cogs Value",
            "Avg. Cost",
            "Selling Price",
            "Inventory Acc.",
            "Cogs. Acc (Default)",
            "Income Acc.(Default)",
        ]

        lines = []
        for pid, d in agg.items():
            product = request.env["product.product"].browse(pid)
            uom = product.uom_id.name if product else ""

            net_in = d["purchased"] + d["gains"] + d["return_cust"]
            net_out = d["sold"] + d["losses"] + d["return_sup"]

            forecast = d["initial"] + net_in + net_out

            avg_cost = abs(round(d["cogs_value"] / d["sold"], 2)) if d["sold"] else (
                product.standard_price if product else 0.0
            )
            forecast_value = forecast * avg_cost
            # avg_standard_cost = abs(round(d["stock_valuation"] / d["purchased"], 2)) if d["purchased"] else (
            #     product.standard_price if product else 0.0
            # )

            lines.append([
                d["product"], d["category"], uom, d["status"],
                d["initial"],
                net_in, net_out,
                d["sold"],
                forecast,
                d["stock_valuation"],
                # forecast_value,
                d["cogs_value"],
                avg_cost,
                d["avg_sales_price"],
                d["stock_valuation_account_id"],
                d["stock_cogs_account_id"],
                d["stock_income_account_id"],

                # d["gains"], d["return_sup"],
                # d["losses"], d["return_cust"],
                # d["purchased"]
            ])

        return headers, lines

    def get_avg_selling_price(self, product, date_from, date_to, excluded_pos_config_ids=None):
        """
        :return: float (average selling price)
        """
        excluded_pos_config_ids = excluded_pos_config_ids or []
        query = """
            SELECT
                CASE
                    WHEN COALESCE(SUM(qty), 0) = 0 THEN 0
                    ELSE COALESCE(SUM(revenue), 0) / SUM(qty)
                END AS avg_price
            FROM (
                -- Sale Orders
                SELECT
                    SUM(sol.price_subtotal) AS revenue,
                    SUM(sol.product_uom_qty) AS qty
                FROM sale_order_line sol
                JOIN sale_order so ON so.id = sol.order_id
                WHERE sol.product_id = %s
                  AND so.date_order BETWEEN %s AND %s
                  AND so.state IN ('sale', 'done')
                  AND sol.product_uom_qty > 0

                UNION ALL

                -- POS Orders (exclude specific POS configs)
                SELECT
                    SUM(pol.price_subtotal) AS revenue,
                    SUM(pol.qty) AS qty
                FROM pos_order_line pol
                JOIN pos_order po ON po.id = pol.order_id
                WHERE pol.product_id = %s
                  AND po.date_order BETWEEN %s AND %s
                  AND po.state IN ('paid', 'done', 'invoiced')
                  AND pol.qty > 0
                  AND po.config_id != ALL(%s)
            ) combined
        """

        request.env.cr.execute(query, (product.id, date_from, date_to, product.id, date_from, date_to, excluded_pos_config_ids))
        row = request.env.cr.fetchone()
        return float(row[0]) if row and row[0] is not None else 0.0
