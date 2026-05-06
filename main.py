import pandas as pd
import os
from telegram import Update
from telegram.ext import ApplicationBuilder, CommandHandler, ContextTypes

# ==========================================
# 1. 配置區域
# ==========================================
TOKEN = os.getenv('TELEGRAM_TOKEN', '8596950787:AAH59ObyGWDz76_VW3Zx3JA1dH9cGsitjaU')
EXCEL_FILE = 'CardShop_System.xlsx'
BRANCH_LIST = ['MK', 'KT', 'TP', 'WAREHOUSE', 'HOME']

# ==========================================
# 2. Excel 核心函數
# ==========================================
def init_excel():
    if not os.path.exists(EXCEL_FILE):
        with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl') as writer:
            pd.DataFrame(columns=['Product', 'Batch', 'Qty', 'Cost', 'Location', 'Time']).to_excel(writer, sheet_name='Inventory', index=False)
            pd.DataFrame(columns=['Time', 'Product', 'Batch', 'Qty', 'Price', 'Cost', 'Branch', 'Profit']).to_excel(writer, sheet_name='Sales', index=False)
        print("📁 已建立全新的 Excel 結構")

def load_inventory():
    if not os.path.exists(EXCEL_FILE):
        init_excel()
    return pd.read_excel(EXCEL_FILE, sheet_name='Inventory')

def load_sales():
    if not os.path.exists(EXCEL_FILE):
        init_excel()
    return pd.read_excel(EXCEL_FILE, sheet_name='Sales')

def save_data(inv_df: pd.DataFrame, sales_df: pd.DataFrame):
    with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl', mode='w') as writer:
        inv_df.to_excel(writer, sheet_name='Inventory', index=False)
        sales_df.to_excel(writer, sheet_name='Sales', index=False)

init_excel()

def is_valid_branch(branch_name: str) -> bool:
    return branch_name.upper() in BRANCH_LIST

# ==========================================
# 3. 指令邏輯
# ==========================================

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    msg = (
        "🚀 卡舖會計系統已啟動！\n\n"
        "📥 /in [貨名] [批次] [數量] [成本]\n"
        "📦 /batchin     ← 批量入貨\n"
        "📤 /out [貨名] [數量] [售價] [分店]\n"
        "📦 /batchout [分店] ← 批量賣貨\n"
        "🔍 /check [分店]\n"
        "💰 /sales       ← 銷售總結 + 日期範圍\n"
        "🚚 /move [貨名] [批次] [數量] [來源] [目的]\n"
        "🗑️ /undo\n\n"
        "💡 分店：MK / KT / TP / WAREHOUSE / HOME"
    )
    await update.message.reply_text(msg)

# /sales - 銷售總結 + 日期範圍查詢
async def handle_sales(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        sales_df = load_sales()
        if sales_df.empty:
            return await update.message.reply_text("📭 暫無任何銷售紀錄")

        sales_df['Date'] = pd.to_datetime(sales_df['Time']).dt.strftime('%Y-%m-%d')
        args = context.args

        if len(args) == 0:
            total_sales = (sales_df['Price'] * sales_df['Qty']).sum()
            total_profit = sales_df['Profit'].sum()
            today = pd.Timestamp.now().strftime('%Y-%m-%d')
            this_month = pd.Timestamp.now().strftime('%Y-%m')

            today_df = sales_df[sales_df['Date'] == today]
            month_df = sales_df[sales_df['Date'].str.startswith(this_month)]

            msg = (
                f"💰 **銷售總結**\n\n"
                f"📊 總生意額：${total_sales:,.0f}\n"
                f"📈 總利潤：${total_profit:,.0f}\n"
                f"📅 今日生意：${(today_df['Price'] * today_df['Qty']).sum():,.0f}\n"
                f"📅 本月生意：${(month_df['Price'] * month_df['Qty']).sum():,.0f}\n\n"
                f"共 {len(sales_df)} 筆交易"
            )
            await update.message.reply_text(msg)

        elif len(args) == 2:
            start_str = args[0].replace('-', '')
            end_str = args[1].replace('-', '')
            mask = (sales_df['Date'] >= f"{start_str[:4]}-{start_str[4:6]}-{start_str[6:]}") & \
                   (sales_df['Date'] <= f"{end_str[:4]}-{end_str[4:6]}-{end_str[6:]}")
            period_df = sales_df[mask]

            if period_df.empty:
                return await update.message.reply_text(f"📭 {start_str} 到 {end_str} 期間冇銷售紀錄")

            period_sales = (period_df['Price'] * period_df['Qty']).sum()
            period_profit = period_df['Profit'].sum()

            msg = (
                f"📅 **銷售紀錄查詢**\n"
                f"🗓️ 日期：{start_str} ~ {end_str}\n\n"
                f"💰 總生意額：${period_sales:,.0f}\n"
                f"📈 總利潤：${period_profit:,.0f}\n"
                f"📦 交易筆數：{len(period_df)} 筆\n\n"
                f"按分店統計："
            )
            for branch, group in period_df.groupby('Branch'):
                b_sales = (group['Price'] * group['Qty']).sum()
                b_profit = group['Profit'].sum()
                msg += f"\n📍 {branch}: ${b_sales:,.0f} (利潤 ${b_profit:,.0f})"
            await update.message.reply_text(msg)

        else:
            await update.message.reply_text("❌ 格式錯誤！\n正確用法：\n/sales\n或\n/sales 2026-05-01 2026-05-07")
    except Exception as e:
        await update.message.reply_text(f"❌ 查詢錯誤: {str(e)}")

# /in - 入貨
async def handle_in(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        if len(context.args) < 4:
            return await update.message.reply_text("❌ 格式錯誤！\n正確格式：/in [貨名] [批次] [數量] [成本]")
        prod = str(context.args[0]).strip()
        batch = str(context.args[1]).strip()
        qty = int(context.args[2])
        cost = float(context.args[3])
        if qty <= 0 or cost <= 0:
            return await update.message.reply_text("❌ 數量同成本必須大過 0！")
        inv_df = load_inventory()
        sales_df = load_sales()
        new_row = {'Product': prod, 'Batch': batch, 'Qty': qty, 'Cost': cost,
                   'Location': 'WAREHOUSE', 'Time': pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')}
        inv_df = pd.concat([inv_df, pd.DataFrame([new_row])], ignore_index=True)
        save_data(inv_df, sales_df)
        await update.message.reply_text(f"✅ 成功入庫：{prod} ({batch}) {qty}盒 @ WAREHOUSE")
    except Exception as e:
        await update.message.reply_text(f"❌ 系統錯誤: {str(e)}")

# /batchin - 批量入貨
async def handle_batchin(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        lines = update.message.text.strip().splitlines()
        if len(lines) < 2:
            return await update.message.reply_text("❌ 格式錯誤！\n正確用法：/batchin\n[貨名1] [批次1] [數量1] [成本1]")
        inv_df = load_inventory()
        sales_df = load_sales()
        success_list = []
        for line in lines[1:]:
            args = [a.strip() for a in line.split() if a.strip()]
            if len(args) < 4: continue
            prod = str(args[0])
            batch = str(args[1])
            qty = int(args[2])
            cost = float(args[3])
            if qty <= 0 or cost <= 0: continue
            new_row = {'Product': prod, 'Batch': batch, 'Qty': qty, 'Cost': cost,
                       'Location': 'WAREHOUSE', 'Time': pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')}
            inv_df = pd.concat([inv_df, pd.DataFrame([new_row])], ignore_index=True)
            success_list.append(f"✅ {prod} ({batch}) {qty}盒")
        save_data(inv_df, sales_df)
        await update.message.reply_text("📦 批量入貨完成！\n\n" + "\n".join(success_list))
    except Exception as e:
        await update.message.reply_text(f"❌ 批量入貨錯誤: {str(e)}")

# /out - 單一賣貨 (FIFO)
async def handle_out(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        args = [a.strip() for a in context.args if a.strip()]
        if len(args) < 4:
            return await update.message.reply_text("❌ 格式錯誤！\n正確格式：/out [貨名] [數量] [售價] [分店]")
        search_name = args[0]
        qty_to_sell = int(args[1])
        price = float(args[2])
        branch = args[3].upper()
        if not is_valid_branch(branch):
            return await update.message.reply_text(f"❌ 錯誤分店：{branch}")
        inv_df = load_inventory()
        sales_df = load_sales()
        mask = (inv_df['Product'].astype(str) == search_name) & \
               (inv_df['Location'].str.upper() == branch) & \
               (inv_df['Qty'] > 0)
        available = inv_df[mask].sort_values('Time')
        if available.empty:
            return await update.message.reply_text(f"❌ {branch} 冇 {search_name} 存貨")
        remaining = qty_to_sell
        now = pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')
        sold_details = []
        for _, row in available.iterrows():
            if remaining <= 0: break
            idx = row.name
            take = min(row['Qty'], remaining)
            profit = (price - row['Cost']) * take
            inv_df.at[idx, 'Qty'] -= take
            new_sale = {'Time': now, 'Product': search_name, 'Batch': row['Batch'],
                        'Qty': take, 'Price': price, 'Cost': row['Cost'],
                        'Branch': branch, 'Profit': profit}
            sales_df = pd.concat([sales_df, pd.DataFrame([new_sale])], ignore_index=True)
            remaining -= take
            sold_details.append(f"{row['Batch']}(x{int(take)})")
        if remaining > 0:
            return await update.message.reply_text(f"⚠️ {branch} 存貨唔夠，仲差 {int(remaining)} 盒！")
        save_data(inv_df, sales_df)
        await update.message.reply_text(
            f"💰 賣貨成功！\n"
            f"📍 分店：{branch}\n"
            f"📦 產品：{search_name}\n"
            f"📆 批次：{' + '.join(sold_details)}"
        )
    except Exception as e:
        await update.message.reply_text(f"❌ 系統錯誤: {str(e)}")

# /batchout - 批量賣貨
async def handle_batchout(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        lines = update.message.text.strip().splitlines()
        if len(lines) < 2:
            return await update.message.reply_text("❌ 格式錯誤！\n正確用法：/batchout [分店]\n[貨名1] [數量1] [售價1]")
        first = lines[0].strip().split()
        branch = first[1].upper() if len(first) > 1 else ""
        if not is_valid_branch(branch):
            return await update.message.reply_text(f"❌ 錯誤分店：{branch}")
        inv_df = load_inventory()
        sales_df = load_sales()
        now = pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')
        total_success = []
        total_profit = 0
        for line in lines[1:]:
            args = [a.strip() for a in line.split() if a.strip()]
            if len(args) < 3: continue
            product = args[0]
            qty_to_sell = int(args[1])
            price = float(args[2])
            mask = (inv_df['Product'].astype(str) == product) & \
                   (inv_df['Location'].str.upper() == branch) & \
                   (inv_df['Qty'] > 0)
            available = inv_df[mask].sort_values('Time')
            if available.empty:
                total_success.append(f"❌ {product} 冇存貨")
                continue
            remaining = qty_to_sell
            sold_details = []
            for _, row in available.iterrows():
                if remaining <= 0: break
                idx = row.name
                take = min(row['Qty'], remaining)
                profit = (price - row['Cost']) * take
                inv_df.at[idx, 'Qty'] -= take
                new_sale = {'Time': now, 'Product': product, 'Batch': row['Batch'],
                            'Qty': take, 'Price': price, 'Cost': row['Cost'],
                            'Branch': branch, 'Profit': profit}
                sales_df = pd.concat([sales_df, pd.DataFrame([new_sale])], ignore_index=True)
                remaining -= take
                sold_details.append(f"{row['Batch']}(x{int(take)})")
                total_profit += profit
            if remaining > 0:
                total_success.append(f"⚠️ {product} 只賣到 {qty_to_sell-remaining} 盒")
            else:
                total_success.append(f"✅ {product} {qty_to_sell}盒 → {' + '.join(sold_details)}")
        save_data(inv_df, sales_df)
        result = f"📦 批量銷售完成！\n📍 分店：{branch}\n\n" + "\n".join(total_success)
        if total_profit > 0:
            result += f"\n\n💰 總利潤：${total_profit:,.0f}"
        await update.message.reply_text(result)
    except Exception as e:
        await update.message.reply_text(f"❌ 批量銷售錯誤: {str(e)}")

# /check - 查庫存（顯示成本）
async def handle_check(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        inv_df = load_inventory()
        df = inv_df[inv_df['Qty'] > 0].copy()
        if df.empty:
            return await update.message.reply_text("📭 全公司目前無任何存貨")
        if context.args:
            target = context.args[0].upper().strip()
            if target not in BRANCH_LIST:
                return await update.message.reply_text(f"❌ 無效分店：{target}")
            branch_df = df[df['Location'].str.upper() == target]
            if branch_df.empty:
                return await update.message.reply_text(f"📍 {target} 目前冇貨")
            report = f"📋 {target} 庫存清單：\n"
            for _, row in branch_df.iterrows():
                report += f" • {row['Product']} ({row['Batch']}): {int(row['Qty'])} 盒 (成本 ${row['Cost']:.0f})\n"
        else:
            report = "🏢 全公司各分店庫存：\n"
            for branch in BRANCH_LIST:
                branch_df = df[df['Location'].str.upper() == branch]
                if not branch_df.empty:
                    report += f"\n📍 【{branch}】\n"
                    for _, row in branch_df.iterrows():
                        report += f"  - {row['Product']} ({row['Batch']}): {int(row['Qty'])} 盒 (成本 ${row['Cost']:.0f})\n"
        await update.message.reply_text(report)
    except Exception as e:
        await update.message.reply_text(f"❌ 查詢錯誤: {str(e)}")

# /report - 今日報表
async def handle_report(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        sales_df = load_sales()
        if sales_df.empty:
            return await update.message.reply_text("📅 暫無任何銷售紀錄")
        today = pd.Timestamp.now().strftime('%Y-%m-%d')
        today_sales = sales_df[sales_df['Time'].astype(str).str.startswith(today)]
        msg = f"💰 今日報表 ({today})\n"
        if today_sales.empty:
            msg += " (今日暫無銷售)"
        else:
            summary = today_sales.groupby('Branch').agg({'Profit': 'sum', 'Qty': 'sum'}).round(0)
            for b, row in summary.iterrows():
                msg += f"📍 {b}: 利潤 ${row['Profit']:,.0f} | 賣出 {int(row['Qty'])} 盒\n"
        await update.message.reply_text(msg)
    except Exception as e:
        await update.message.reply_text(f"❌ 報表錯誤: {str(e)}")

# /move - 調貨
async def handle_move(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        if len(context.args) < 5:
            return await update.message.reply_text("❌ 格式錯誤！\n正確格式：/move [貨名] [批次] [數量] [來源] [目的]")
        prod = str(context.args[0]).strip()
        batch = str(context.args[1]).strip()
        qty = int(context.args[2])
        from_loc = str(context.args[3]).strip().upper()
        to_loc = str(context.args[4]).strip().upper()
        if not is_valid_branch(from_loc) or not is_valid_branch(to_loc):
            return await update.message.reply_text("❌ 來源或目的地分店無效")
        if from_loc == to_loc:
            return await update.message.reply_text("❌ 來源同目的地不能一樣")
        inv_df = load_inventory()
        sales_df = load_sales()
        inv_df['Product_Str'] = inv_df['Product'].astype(str).str.strip()
        inv_df['Batch_Str'] = inv_df['Batch'].astype(str).str.strip()
        inv_df['Loc_Str'] = inv_df['Location'].astype(str).str.strip().str.upper()
        mask = (inv_df['Product_Str'] == prod) & (inv_df['Batch_Str'] == batch) & (inv_df['Loc_Str'] == from_loc)
        match_idx = inv_df[mask].index
        if match_idx.empty:
            avail = inv_df[(inv_df['Product_Str'] == prod) & (inv_df['Loc_Str'] == from_loc) & (inv_df['Qty'] > 0)]
            if avail.empty:
                return await update.message.reply_text(f"❌ {from_loc} 完全冇 {prod} 存貨")
            else:
                msg = f"❌ {from_loc} 搵唔到 {prod} (batch: {batch})\n\n📋 目前可用批次：\n"
                for _, row in avail.iterrows():
                    msg += f"   • batch {row['Batch_Str']}：{int(row['Qty'])}盒 (成本 ${row['Cost']:.0f})\n"
                return await update.message.reply_text(msg)
        current_stock = inv_df.at[match_idx[0], 'Qty']
        if current_stock < qty:
            return await update.message.reply_text(f"❌ {from_loc} 貨量不足，只剩 {int(current_stock)} 盒")
        cost = inv_df.at[match_idx[0], 'Cost']
        inv_df.at[match_idx[0], 'Qty'] -= qty
        dest_mask = (inv_df['Product_Str'] == prod) & (inv_df['Batch_Str'] == batch) & (inv_df['Loc_Str'] == to_loc)
        dest_idx = inv_df[dest_mask].index
        if not dest_idx.empty:
            inv_df.at[dest_idx[0], 'Qty'] += qty
        else:
            new_row = {'Product': prod, 'Batch': batch, 'Qty': qty, 'Cost': cost,
                       'Location': to_loc, 'Time': pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')}
            inv_df = pd.concat([inv_df, pd.DataFrame([new_row])], ignore_index=True)
        inv_df = inv_df.drop(columns=['Product_Str', 'Batch_Str', 'Loc_Str'])
        save_data(inv_df, sales_df)
        await update.message.reply_text(f"🚚 調貨成功！\n📦 {prod} ({batch})\n📍 {from_loc} ➡️ {to_loc}\n🔢 數量：{qty}盒")
    except Exception as e:
        await update.message.reply_text(f"❌ 調貨錯誤: {str(e)}")

# /undo - 撤銷上次賣貨
async def handle_undo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        sales_df = load_sales()
        inv_df = load_inventory()
        if sales_df.empty:
            return await update.message.reply_text("📭 冇任何銷售紀錄可以撤銷")
        last_sale = sales_df.iloc[-1].copy()
        product = last_sale['Product']
        batch = last_sale['Batch']
        qty = int(last_sale['Qty'])
        branch = last_sale['Branch']
        cost = last_sale['Cost']
        mask = (inv_df['Product'].astype(str) == product) & \
               (inv_df['Batch'].astype(str) == batch) & \
               (inv_df['Location'].str.upper() == branch)
        if mask.any():
            idx = inv_df[mask].index[0]
            inv_df.at[idx, 'Qty'] += qty
        else:
            new_row = {'Product': product, 'Batch': batch, 'Qty': qty, 'Cost': cost,
                       'Location': branch, 'Time': pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')}
            inv_df = pd.concat([inv_df, pd.DataFrame([new_row])], ignore_index=True)
        sales_df = sales_df.iloc[:-1].reset_index(drop=True)
        save_data(inv_df, sales_df)
        await update.message.reply_text(f"🗑️ 已撤銷！\n📦 {product} ({batch}) x{qty} @ {branch}")
    except Exception as e:
        await update.message.reply_text(f"❌ 撤銷失敗: {str(e)}")

# ==========================================
# 4. 主程式
# ==========================================
if __name__ == '__main__':
    app = ApplicationBuilder().token(TOKEN).build()
    
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("in", handle_in))
    app.add_handler(CommandHandler("batchin", handle_batchin))
    app.add_handler(CommandHandler("out", handle_out))
    app.add_handler(CommandHandler("batchout", handle_batchout))
    app.add_handler(CommandHandler("check", handle_check))
    app.add_handler(CommandHandler("sales", handle_sales))
    app.add_handler(CommandHandler("report", handle_report))
    app.add_handler(CommandHandler("move", handle_move))
    app.add_handler(CommandHandler("undo", handle_undo))

    app.bot.set_my_commands([])

    print("🚀 卡舖會計系統（最完整版 + /sales 日期範圍）已上線！")
    app.run_polling()