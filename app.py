import streamlit as st
import pandas as pd
from datetime import datetime
import io
import os
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# PDF生成用（ReportLab）
from reportlab.lib.pagesizes import A4, portrait
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# 履歴保存・設定用のファイル名
HISTORY_FILE = "sales_history.csv"
EXCEL_FILE = "maker_price_list.xlsx"
COMPANY_PRESET_FILE = "company_presets.csv"
BANK_PRESET_FILE = "bank_presets.csv"

# 日本語フォントの設定（安全なフォールバック付き）
FONT_NAME = 'Helvetica'
try:
    if os.path.exists("C:/Windows/Fonts/meiryo.ttc"):
        pdfmetrics.registerFont(TTFont('Meiryo', 'C:/Windows/Fonts/meiryo.ttc'))
        FONT_NAME = 'Meiryo'
    elif os.path.exists("C:/Windows/Fonts/msgothic.ttc"):
        pdfmetrics.registerFont(TTFont('Msgothic', 'C:/Windows/Fonts/msgothic.ttc'))
        FONT_NAME = 'Msgothic'
except Exception:
    pass

# ページの設定
st.set_page_config(page_title="アパレルプリント受発注システム", layout="centered")

# ==========================================
# 🗂️ サイドバーによる画面（ページ）切り替え
# ==========================================
st.sidebar.title("📌 メニュー切り替え")
app_mode = st.sidebar.radio(
    "移動先を選択してください",
    [
        "👕 1. ボディ発注（カート・注文書作成）", 
        "📦 2. 在庫リスト確認・編集",
        "💰 3. 見積書・売価設定・履歴", 
        "📊 4. 粗利計算・利益管理",
        "📄 5. 請求書作成・発行"
    ]
)

# 共通のエクセル行位置設定
NAME_ROW_IDX = 2      
ITEM_ROW_IDX = 1      
COLOR_ROW_IDX = 4     
SIZE_START_ROW = 5    

# セッションステートの初期化
if "cart_items" not in st.session_state:
    st.session_state.cart_items = []
if "confirmed_stock" not in st.session_state:
    st.session_state.confirmed_stock = []
if "order_completed" not in st.session_state:
    st.session_state.order_completed = False
if "last_ordered_items" not in st.session_state:
    st.session_state["last_ordered_items"] = []

# デフォルトのプリセットCSV作成（初回のみ）
def init_presets():
    if not os.path.exists(COMPANY_PRESET_FILE):
        df_c = pd.DataFrame([{
            "名称": "デフォルト自社",
            "自社名": "〇〇株式会社",
            "自社住所": "〒123-4567 〇〇県〇〇市...",
            "自社担当": "担当 太郎"
        }])
        df_c.to_csv(COMPANY_PRESET_FILE, index=False, encoding="utf-8-sig")
        
    if not os.path.exists(BANK_PRESET_FILE):
        df_b = pd.DataFrame([{
            "名称": "メイン口座",
            "銀行名": "〇〇銀行",
            "支店名": "〇〇支店",
            "口座番号": "普通 1234567",
            "口座名義": "カ）〇〇カンパニー"
        }])
        df_b.to_csv(BANK_PRESET_FILE, index=False, encoding="utf-8-sig")

init_presets()

# 単価検索関数
def search_unit_price(excel_path: str, sheet_name: str, target_item: str, target_color: str, target_size: str):
    df = pd.read_excel(excel_path, sheet_name=sheet_name, header=None)
    item_row = df.iloc[ITEM_ROW_IDX].values
    color_row = df.iloc[COLOR_ROW_IDX].values
    
    white_keywords = ["ホワイト", "白", "white", "WHITE"]
    is_target_white = target_color.strip() in white_keywords
    
    matched_col_index = None
    for col_idx in range(len(item_row)):
        if str(item_row[col_idx]).strip() == target_item:
            col_color = str(color_row[col_idx]).strip()
            is_col_white = col_color in white_keywords
            
            if is_target_white and is_col_white:
                matched_col_index = col_idx
                break
            elif not is_target_white and not is_col_white and col_color != "nan" and col_color != "":
                matched_col_index = col_idx
                break
            
    if matched_col_index is None:
        raise ValueError(f"品番 '{target_item}' とカラー '{target_color}' の価格列が見つかりませんでした。")

    matched_row_index = None
    for row_idx in range(SIZE_START_ROW, len(df)):
        s1 = str(df.iloc[row_idx, 0]).strip()
        s2 = str(df.iloc[row_idx, 1]).strip()
        if s1 == str(target_size) or s2 == str(target_size):
            matched_row_index = row_idx
            break
            
    if matched_row_index is None:
        raise ValueError(f"サイズ '{target_size}' が見つかりませんでした。")

    return int(df.iloc[matched_row_index, matched_col_index])

# ==========================================
# 画面 1：ボディ発注（カート・注文書作成）
# ==========================================
if app_mode == "👕 1. ボディ発注（カート・注文書作成）":
    st.title("👕 ボディ発注書作成（カート方式）")
    st.write("メーカーの価格リストからボディを選び、一度カートに入れてから注文を確定・発注書を発行します。")

    maker_input = st.selectbox("メーカー選択", ["toms", "cab"])
    product_name = ""
    selected_item = ""
    selected_size = ""

    try:
        sheet_name = f"{maker_input}価格リスト"
        df_master = pd.read_excel(EXCEL_FILE, sheet_name=sheet_name, header=None)
        
        item_row = df_master.iloc[ITEM_ROW_IDX].values[2:]
        available_items = []
        for val in item_row:
            val_str = str(val).strip()
            if val_str and val_str != "nan" and val_str not in available_items:
                available_items.append(val_str)
                
        selected_item = st.selectbox("品番選択", available_items)

        for col_idx in range(2, df_master.shape[1]):
            if str(df_master.iloc[ITEM_ROW_IDX, col_idx]).strip() == selected_item:
                name_val = str(df_master.iloc[NAME_ROW_IDX, col_idx]).strip()
                if name_val and name_val != "nan":
                    product_name = name_val
                    break
        
        if product_name:
            st.info(f"🏷️ **商品名称:** {product_name}")

        item_cols = []
        for col_idx in range(2, df_master.shape[1]):
            if str(df_master.iloc[ITEM_ROW_IDX, col_idx]).strip() == selected_item:
                item_cols.append(col_idx)

        available_sizes = []
        for row_idx in range(SIZE_START_ROW, len(df_master)):
            s1 = str(df_master.iloc[row_idx, 0]).strip()
            s2 = str(df_master.iloc[row_idx, 1]).strip()
            size_val = s1 if s1 and s1 != "nan" else s2
            
            if not size_val or size_val == "nan":
                continue
                
            has_price = False
            for c_idx in item_cols:
                price_val = df_master.iloc[row_idx, c_idx]
                if pd.notna(price_val) and str(price_val).strip() != "" and str(price_val).strip() != "nan":
                    has_price = True
                    break
            
            if has_price and size_val not in available_sizes:
                available_sizes.append(size_val)
                
        selected_size = st.selectbox("サイズ選択", available_sizes)

    except Exception as e:
        st.error(f"価格表の読み込みに失敗しました: {e}")

    color_input = st.text_input("カラー入力（例: ホワイト、ブラックなど）", value="ブラック")
    purchase_qty_input = st.number_input("追加枚数", min_value=1, value=1)

    if st.button("🛒 カートに追加する"):
        cleaned_color = color_input.strip()
        
        found = False
        for item in st.session_state.cart_items:
            if (item["maker"] == maker_input and
                item["item_code"] == selected_item and
                item["color"] == cleaned_color and
                item["size"] == selected_size):
                item["purchase_quantity"] += purchase_qty_input
                found = True
                break
        
        if not found:
            st.session_state.cart_items.append({
                "maker": maker_input,
                "item_code": selected_item,
                "product_name": product_name,
                "color": cleaned_color,
                "size": selected_size,
                "purchase_quantity": purchase_qty_input
            })
            
        st.success("カートに商品を追加しました！")

    if st.session_state.cart_items:
        st.write("---")
        st.subheader("🛒 現在のカート内アイテム・小計")
        
        total_cart_price = 0
        for idx, item in enumerate(st.session_state.cart_items):
            try:
                unit_p = search_unit_price(
                    excel_path=EXCEL_FILE,
                    sheet_name=f"{item['maker']}価格リスト",
                    target_item=item["item_code"],
                    target_color=item["color"],
                    target_size=item["size"]
                )
            except Exception:
                unit_p = 0
            
            sub_total = unit_p * item["purchase_quantity"]
            total_cart_price += sub_total

            st.markdown(
                f"**{idx + 1}. [{item['maker'].upper()}] {item['item_code']} ({item['product_name']})**<br/>"
                f"カラー: {item['color']} / サイズ: {item['size']} / "
                f"単価: **{unit_p:,}円** × 枚数: **{item['purchase_quantity']}枚** = 小計: **{sub_total:,}円**",
                unsafe_allow_html=True
            )
            st.write("")

        st.markdown(f"### 🧮 カート合計金額 (税抜): **{total_cart_price:,} 円**")
        st.write("---")

        if st.button("🗑️ カートを空にする"):
            st.session_state.cart_items = []
            st.session_state.order_completed = False
            st.session_state["last_ordered_items"] = []
            st.rerun()

        st.write("")
        if st.button("✅ 【注文確定】して在庫リストに追加 ＆ 注文書を有効化する"):
            for cart_item in st.session_state.cart_items:
                stock_found = False
                for stock_item in st.session_state.confirmed_stock:
                    if (stock_item["maker"] == cart_item["maker"] and
                        stock_item["item_code"] == cart_item["item_code"] and
                        stock_item["color"] == cart_item["color"] and
                        stock_item["size"] == cart_item["size"]):
                        stock_item["purchase_quantity"] += cart_item["purchase_quantity"]
                        stock_found = True
                        break
                if not stock_found:
                    st.session_state.confirmed_stock.append(cart_item.copy())
            
            st.session_state["last_ordered_items"] = st.session_state.cart_items.copy()
            st.session_state.order_completed = True
            st.session_state.cart_items = []
            st.success("🎉 注文が確定し、在庫リストへ反映されました！")
            st.rerun()

    if st.session_state.order_completed and st.session_state["last_ordered_items"]:
        st.write("---")
        st.subheader("📥 直近の注文書ダウンロード")
        
        col_dl1, col_dl2 = st.columns(2)
        with col_dl1:
            try:
                export_rows = []
                for item in st.session_state["last_ordered_items"]:
                    maker = item["maker"]
                    sheet_name = f"{maker}価格リスト"
                    purchase_qty = item["purchase_quantity"]
                    unit_price = search_unit_price(EXCEL_FILE, sheet_name, item["item_code"], item["color"], item["size"])
                    export_rows.append({
                        "メーカー": maker.upper(),
                        "品番": item["item_code"],
                        "商品名称": item.get("product_name", ""),
                        "カラー": item["color"],
                        "サイズ": item["size"],
                        "ボディ単価": unit_price,
                        "発注数": purchase_qty,
                        "ボディ小計": unit_price * purchase_qty
                    })
                result_df = pd.DataFrame(export_rows)
                output_excel = io.BytesIO()
                with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
                    result_df.to_excel(writer, sheet_name="ボディ注文書", index=False)
                st.download_button(
                    label="📥 ボディ発注書Excelを保存",
                    data=output_excel.getvalue(),
                    file_name=f"ボディ発注書_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"Excel生成エラー: {e}")

        with col_dl2:
            try:
                pdf_buffer = io.BytesIO()
                doc = SimpleDocTemplate(pdf_buffer, pagesize=portrait(A4), rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
                
                t_style = ParagraphStyle('OrdTitle', fontName=FONT_NAME, fontSize=16, textColor=colors.white, alignment=0, spaceBefore=4, spaceAfter=4)
                right_meta_s = ParagraphStyle('OrdRMeta', fontName=FONT_NAME, fontSize=9, alignment=2)
                meta_s = ParagraphStyle('OrdMeta', fontName=FONT_NAME, fontSize=9)
                
                cell_s = ParagraphStyle('OrdCell', fontName=FONT_NAME, fontSize=8)
                cell_c = ParagraphStyle('OrdCellC', fontName=FONT_NAME, fontSize=8, alignment=1)
                cell_r = ParagraphStyle('OrdCellR', fontName=FONT_NAME, fontSize=8, alignment=2)
                hdr_s = ParagraphStyle('OrdHdr', fontName=FONT_NAME, fontSize=8, textColor=colors.white, alignment=1)
                
                sum_hdr_style = ParagraphStyle('SumHdr', fontName=FONT_NAME, fontSize=8, textColor=colors.white, alignment=1)
                sum_val_style = ParagraphStyle('SumVal', fontName=FONT_NAME, fontSize=8, textColor=colors.black, alignment=2)

                elements = []

                header_banner_data = [
                    [Paragraph("<b>ボ デ ィ 発 注 書</b>", t_style), Paragraph(f"発行日 : {datetime.now().strftime('%Y年%m月%d日')}", right_meta_s)]
                ]
                t_banner = Table(header_banner_data, colWidths=[300, 235])
                t_banner.setStyle(TableStyle([
                    ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#2E6EA5')),
                    ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                    ('BOTTOMPADDING', (0,0), (-1,-1), 8),
                    ('TOPPADDING', (0,0), (-1,-1), 8),
                ]))
                elements.append(t_banner)
                elements.append(Spacer(1, 15))
                elements.append(Paragraph("メーカー各社ご担当者様<br/>下記のとおり発注いたします。", meta_s))
                elements.append(Spacer(1, 10))

                table_data = [
                    [Paragraph("メーカー", hdr_s), Paragraph("品番 ／ 商品名称", hdr_s), Paragraph("カラー / サイズ", hdr_s), Paragraph("単価", hdr_s), Paragraph("数量", hdr_s), Paragraph("小計", hdr_s)]
                ]

                total_order_price = 0
                total_order_qty = 0

                for item in st.session_state["last_ordered_items"]:
                    maker = item["maker"]
                    sheet_name = f"{maker}価格リスト"
                    p_qty = item["purchase_quantity"]
                    try:
                        u_price = search_unit_price(EXCEL_FILE, sheet_name, item["item_code"], item["color"], item["size"])
                    except Exception:
                        u_price = 0
                    
                    sub_p = u_price * p_qty
                    total_order_price += sub_p
                    total_order_qty += p_qty

                    table_data.append([
                        Paragraph(maker.upper(), cell_c),
                        Paragraph(f"<b>{item['item_code']}</b><br/>{item.get('product_name', '')}", cell_s),
                        Paragraph(f"{item['color']} / {item['size']}", cell_s),
                        Paragraph(f"{u_price:,} 円", cell_r),
                        Paragraph(str(p_qty), cell_c),
                        Paragraph(f"{sub_p:,} 円", cell_r)
                    ])

                for _ in range(max(0, 10 - len(st.session_state["last_ordered_items"]))):
                    table_data.append([Paragraph("", cell_s), Paragraph("", cell_s), Paragraph("", cell_s), Paragraph("", cell_r), Paragraph("", cell_c), Paragraph("", cell_r)])

                t_details = Table(table_data, colWidths=[60, 154, 130, 60, 50, 81])
                ts = [
                    ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#2E6EA5')),
                    ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                    ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#BDC3C7')),
                    ('TOPPADDING', (0,0), (-1,-1), 5),
                    ('BOTTOMPADDING', (0,0), (-1,-1), 5),
                ]
                for r_idx in range(1, len(table_data)):
                    if r_idx % 2 == 1:
                        ts.append(('BACKGROUND', (0, r_idx), (-1, r_idx), colors.HexColor('#F2F4F7')))
                    else:
                        ts.append(('BACKGROUND', (0, r_idx), (-1, r_idx), colors.white))
                
                t_details.setStyle(TableStyle(ts))
                elements.append(t_details)
                elements.append(Spacer(1, 10))

                sum_table_data = [
                    [
                        Paragraph("総発注枚数", sum_hdr_style), Paragraph(f"{total_order_qty:,} 枚", sum_val_style),
                        Paragraph("合計金額 (税抜)", sum_hdr_style), Paragraph(f"¥ {total_order_price:,}", sum_val_style),
                    ]
                ]
                t_sum = Table(sum_table_data, colWidths=[90, 150, 114, 181])
                t_sum.setStyle(TableStyle([
                    ('BACKGROUND', (0,0), (0,0), colors.HexColor('#2E6EA5')),
                    ('BACKGROUND', (2,0), (2,0), colors.HexColor('#2E6EA5')),
                    ('BOX', (0,0), (-1,-1), 0.5, colors.HexColor('#2E6EA5')),
                    ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                    ('TOPPADDING', (0,0), (-1,-1), 5),
                    ('BOTTOMPADDING', (0,0), (-1,-1), 5),
                ]))
                elements.append(t_sum)

                doc.build(elements)

                st.download_button(
                    label="📥 ボディ発注書PDFを保存",
                    data=pdf_buffer.getvalue(),
                    file_name=f"ボディ発注書_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf",
                    mime="application/pdf"
                )
            except Exception as e:
                st.error(f"PDF生成エラー: {e}")

# ==========================================
# 画面 2：在庫リスト確認・編集
# ==========================================
elif app_mode == "📦 2. 在庫リスト確認・編集":
    st.title("📦 確定済み在庫リストの確認・編集")
    if not st.session_state.confirmed_stock:
        st.info("現在、確定している在庫はありません。")
    else:
        for idx, stock_item in enumerate(st.session_state.confirmed_stock):
            with st.expander(f"📦 [{stock_item['maker'].upper()}] 品番: {stock_item['item_code']} - {stock_item['color']} / {stock_item['size']} (現在庫: {stock_item['purchase_quantity']}枚)"):
                col_s1, col_s2, col_s3 = st.columns([2, 1, 1])
                with col_s1:
                    new_qty = st.number_input("在庫数変更", min_value=0, value=int(stock_item["purchase_quantity"]), key=f"stock_edit_qty_{idx}")
                with col_s2:
                    if st.button("💾 数量を更新", key=f"update_stock_{idx}"):
                        st.session_state.confirmed_stock[idx]["purchase_quantity"] = new_qty
                        st.success("更新しました！")
                        st.rerun()
                with col_s3:
                    if st.button("🗑️ この行を削除", key=f"delete_stock_{idx}"):
                        st.session_state.confirmed_stock.pop(idx)
                        st.success("削除しました！")
                        st.rerun()

# ==========================================
# 画面 3：見積書・売価設定・履歴
# ==========================================
elif app_mode == "💰 3. 見積書・売価設定・履歴":
    st.title("💰 見積書・受書作成 ＆ 履歴管理")
    st.write("お見積書を作成・保存すると、下部の履歴一覧からいつでも画像イメージ通りの**「注文請書」**を発行できるようになります。")

    col_input_n1, col_input_n2 = st.columns(2)
    with col_input_n1:
        project_name = st.text_input("案件名", value="〇〇イベント用プリント制作")
    with col_input_n2:
        customer_name = st.text_input("顧客名（お取引先名）", value="株式会社〇〇 御中")

    print_method = st.selectbox("共通プリント方式", ["シルクスクリーン", "インクジェット", "DTF"])
    
    dtf_detail_summary = ""
    dtf_total_cost_calc = 0
    if print_method == "DTF":
        st.markdown("---")
        st.subheader("🖨️ DTFプリント設定")
        
        d_hdr_cols = st.columns([1.5, 1, 1, 1])
        with d_hdr_cols[0]:
            st.markdown("**件名（位置など）**")
        with d_hdr_cols[1]:
            st.markdown("**面付数**")
        with d_hdr_cols[2]:
            st.markdown("**分割数**")
        with d_hdr_cols[3]:
            st.markdown("**使用数**")

        dtf_summary_parts = []
        for d_i in range(3):
            d_cols = st.columns([1.5, 1, 1, 1])
            with d_cols[0]:
                d_title = st.text_input(f"件名 #{d_i+1}", value="前面" if d_i == 0 else "", key=f"dtf_title_{d_i}", label_visibility="collapsed")
            with d_cols[1]:
                d_qty = st.number_input(f"面付 #{d_i+1}", min_value=0, value=1 if d_i == 0 else 0, key=f"dtf_qty_{d_i}", label_visibility="collapsed")
            with d_cols[2]:
                d_split = st.number_input(f"分割 #{d_i+1}", min_value=1, value=1, key=f"dtf_split_{d_i}", label_visibility="collapsed")
            with d_cols[3]:
                d_mult = st.number_input(f"使用数 #{d_i+1}", min_value=1, value=1, key=f"dtf_mult_{d_i}", label_visibility="collapsed")
            
            if d_qty > 0 and d_title.strip() != "":
                l_cost = int(((1400 / d_split) * d_qty) * d_mult)
                dtf_summary_parts.append(f"{d_title}(面付:{d_qty}/分:{d_split}→{l_cost:,}円)")
                dtf_total_cost_calc += l_cost
        if dtf_summary_parts:
            dtf_detail_summary = " / ".join(dtf_summary_parts)

    st.write("---")
    st.subheader("📋 見積明細行の編集（最大15行）")

    stock_options = [
        f"[{item['maker'].upper()}] {item['item_code']} ({item['product_name']}) - {item['color']} / {item['size']} (在庫: {item['purchase_quantity']}枚)"
        for item in st.session_state.confirmed_stock
    ] if st.session_state.confirmed_stock else []

    est_line_items = []
    total_sales_amount = 0
    calculated_body_cost_total = 0

    for i in range(15):
        with st.expander(f"➕ 明細行 {i + 1}", expanded=(i == 0)):
            col_type, col_qty, col_price = st.columns([2, 1, 1])
            with col_type:
                line_mode = st.radio(f"行{i+1}入力元", ["📦 確定在庫から選択", "✍️ 自由入力"], key=f"mode_{i}", horizontal=True)
            
            detail_note = ""
            body_cost = 0
            max_qty_limit = 9999

            if line_mode == "📦 確定在庫から選択" and stock_options:
                selected_stk_idx = st.selectbox(f"在庫選択 #{i+1}", range(len(stock_options)), format_func=lambda x: stock_options[x], key=f"stk_sel_{i}")
                chosen = st.session_state.confirmed_stock[selected_stk_idx]
                detail_note = f"カラー: {chosen['color']} / サイズ: {chosen['size']}"
                max_qty_limit = int(chosen['purchase_quantity'])
                try:
                    body_cost = search_unit_price(EXCEL_FILE, f"{chosen['maker']}価格リスト", chosen["item_code"], chosen["color"], chosen["size"])
                except Exception:
                    body_cost = 0
            else:
                detail_note = st.text_input(f"仕様メモ #{i+1}", value="カラー: ブラック / サイズ: L" if i == 0 else "", key=f"free_note_{i}")
                body_cost = 500

            with col_qty:
                default_qty_val = 1 if (i == 0 and 1 <= max_qty_limit) else 0
                qty = st.number_input(f"数量 #{i+1}", min_value=0, max_value=max_qty_limit, value=default_qty_val, key=f"qty_{i}")
            with col_price:
                selling_price = st.number_input(f"販売単価(税抜) #{i+1}", min_value=0, value=3500 if i == 0 else 0, step=100, key=f"price_{i}")

            if qty > 0 and project_name.strip() != "":
                line_total = qty * selling_price
                total_sales_amount += line_total
                calculated_body_cost_total += (body_cost * qty)
                est_line_items.append({
                    "detail_note": detail_note,
                    "qty": qty,
                    "unit_price": selling_price,
                    "line_total": line_total
                })

    total_qty_sum = sum([item["qty"] for item in est_line_items])
    print_process_cost = st.number_input("プリント加工費 総額 (税抜)", min_value=0, value=dtf_total_cost_calc if print_method == "DTF" else total_qty_sum * 300, step=100)

    st.write("---")
    file_timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')

    if st.button("🚀 履歴に保存して見積書を発行する", disabled=not est_line_items):
        history_df_new = pd.DataFrame([{
            "日時": datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            "案件名": project_name,
            "顧客名": customer_name,
            "品名": project_name,
            "プリント方式": print_method,
            "販売単価": f"{est_line_items[0]['unit_price']:,}円" if est_line_items else "0円",
            "数量": total_qty_sum,
            "売上合計(税抜)": total_sales_amount,
            "自動ボディ原価": calculated_body_cost_total,
            "自動加工費": print_process_cost,
            "自社名": "〇〇株式会社",
            "自社住所": "〒123-4567 〇〇県〇〇市...",
            "自社担当": "担当 太郎",
            "銀行名": "〇〇銀行",
            "支店名": "〇〇支店",
            "口座番号": "普通 1234567",
            "口座名義": "カ）〇〇カンパニー"
        }])
        if os.path.exists(HISTORY_FILE):
            history_df_old = pd.read_csv(HISTORY_FILE)
            for col_n in ["自社名", "自社住所", "自社担当", "銀行名", "支店名", "口座番号", "口座名義"]:
                if col_n not in history_df_old.columns:
                    history_df_old[col_n] = ""
            history_df_combined = pd.concat([history_df_new, history_df_old], ignore_index=True)
        else:
            history_df_combined = history_df_new
        history_df_combined.to_csv(HISTORY_FILE, index=False, encoding="utf-8-sig")
        st.success("見積履歴を保存しました！ 下記から見積書PDFをダウンロードできます。")

    if est_line_items:
        try:
            pdf_buffer = io.BytesIO()
            doc = SimpleDocTemplate(pdf_buffer, pagesize=portrait(A4), rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
            
            t_style = ParagraphStyle('EstTitle', fontName=FONT_NAME, fontSize=16, textColor=colors.white, alignment=0, spaceBefore=4, spaceAfter=4)
            meta_s = ParagraphStyle('EstMeta', fontName=FONT_NAME, fontSize=9)
            right_meta_s = ParagraphStyle('EstRMeta', fontName=FONT_NAME, fontSize=9, alignment=2)
            cust_large_s = ParagraphStyle('EstCustLarge', fontName=FONT_NAME, fontSize=13, leading=16)

            cell_s = ParagraphStyle('EstCell', fontName=FONT_NAME, fontSize=8)
            cell_c = ParagraphStyle('EstCellC', fontName=FONT_NAME, fontSize=8, alignment=1)
            cell_r = ParagraphStyle('EstCellR', fontName=FONT_NAME, fontSize=8, alignment=2)
            hdr_s = ParagraphStyle('EstHdr', fontName=FONT_NAME, fontSize=8, textColor=colors.white, alignment=1)
            
            amt_title_s = ParagraphStyle('EstAmtT', fontName=FONT_NAME, fontSize=11, textColor=colors.white, alignment=1)
            amt_val_s = ParagraphStyle('EstAmtV', fontName=FONT_NAME, fontSize=14, textColor=colors.black, alignment=2)

            sum_hdr_style = ParagraphStyle('EstSumHdr', fontName=FONT_NAME, fontSize=8, textColor=colors.white, alignment=1)
            sum_val_style = ParagraphStyle('EstSumVal', fontName=FONT_NAME, fontSize=8, textColor=colors.black, alignment=2)

            elements = []

            header_banner_data = [
                [Paragraph("<b>御 見 積 書</b>", t_style), Paragraph(f"No : EST-{file_timestamp}<br/>発行日 : {datetime.now().strftime('%Y年%m月%d日')}", right_meta_s)]
            ]
            t_banner = Table(header_banner_data, colWidths=[300, 245])
            t_banner.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#2E6EA5')),
                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                ('BOTTOMPADDING', (0,0), (-1,-1), 8),
                ('TOPPADDING', (0,0), (-1,-1), 8),
            ]))
            elements.append(t_banner)
            elements.append(Spacer(1, 10))

            cust_text = f"<b>{customer_name}</b>"
            company_outside_info = [
                Paragraph("<b>〇〇株式会社</b>", meta_s),
                Paragraph("〒123-4567 〇〇県〇〇市...", meta_s),
                Paragraph("担当 太郎", meta_s),
            ]

            right_column_content = [
                company_outside_info[0],
                company_outside_info[1],
                company_outside_info[2],
            ]

            top_info_data = [
                [Paragraph(cust_text, cust_large_s), right_column_content]
            ]
            t_top = Table(top_info_data, colWidths=[310, 235])
            t_top.setStyle(TableStyle([
                ('VALIGN', (0,0), (-1,-1), 'TOP'),
            ]))
            elements.append(t_top)
            elements.append(Spacer(1, 10))

            elements.append(Paragraph(f"<b>件名：{project_name}</b>", meta_s))
            elements.append(Spacer(1, 4))
            elements.append(Paragraph("下記のとおりお見積り申し上げます。", meta_s))
            elements.append(Spacer(1, 8))

            tax_amount = int(total_sales_amount * 0.1)
            total_inc_tax = total_sales_amount + tax_amount

            amt_box_data = [
                [Paragraph("御見積金額", amt_title_s), Paragraph(f"¥ {total_inc_tax:,} -", amt_val_s)]
            ]
            t_amt = Table(amt_box_data, colWidths=[120, 220])
            t_amt.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (0,0), colors.HexColor('#2E6EA5')),
                ('BOX', (0,0), (-1,-1), 1, colors.HexColor('#2E6EA5')),
                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                ('TOPPADDING', (0,0), (-1,-1), 6),
                ('BOTTOMPADDING', (0,0), (-1,-1), 6),
            ]))
            elements.append(t_amt)
            elements.append(Spacer(1, 12))

            table_data = [
                [Paragraph("摘要 ／ 仕様", hdr_s), Paragraph("単価", hdr_s), Paragraph("数量", hdr_s), Paragraph("金額", hdr_s)]
            ]
            
            for item in est_line_items:
                table_data.append([
                    Paragraph(f"<b>{project_name}</b><br/>{item['detail_note']} ({print_method})", cell_s),
                    Paragraph(f"{item['unit_price']:,} 円", cell_r),
                    Paragraph(str(item['qty']), cell_c),
                    Paragraph(f"{item['line_total']:,} 円", cell_r)
                ])

            for _ in range(max(0, 10 - len(est_line_items))):
                table_data.append([Paragraph("", cell_s), Paragraph("", cell_r), Paragraph("", cell_c), Paragraph("", cell_r)])

            t_details = Table(table_data, colWidths=[264, 80, 60, 141])
            ts = [
                ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#2E6EA5')),
                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#BDC3C7')),
            ]
            for r_idx in range(1, len(table_data)):
                if r_idx % 2 == 1:
                    ts.append(('BACKGROUND', (0, r_idx), (-1, r_idx), colors.HexColor('#F2F4F7')))
                else:
                    ts.append(('BACKGROUND', (0, r_idx), (-1, r_idx), colors.white))
            
            t_details.setStyle(TableStyle(ts))
            elements.append(t_details)
            elements.append(Spacer(1, 8))

            sum_table_data = [
                [
                    Paragraph("小計 (税抜)", sum_hdr_style), Paragraph(f"¥ {total_sales_amount:,}", sum_val_style),
                    Paragraph("消費税 (10%)", sum_hdr_style), Paragraph(f"¥ {tax_amount:,}", sum_val_style),
                    Paragraph("合計 (税込)", sum_hdr_style), Paragraph(f"¥ {total_inc_tax:,}", sum_val_style),
                ]
            ]
            t_sum = Table(sum_table_data, colWidths=[70, 111, 70, 111, 70, 113])
            t_sum.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (0,0), colors.HexColor('#2E6EA5')),
                ('BACKGROUND', (2,0), (2,0), colors.HexColor('#2E6EA5')),
                ('BACKGROUND', (4,0), (4,0), colors.HexColor('#2E6EA5')),
                ('BOX', (0,0), (-1,-1), 0.5, colors.HexColor('#2E6EA5')),
                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                ('TOPPADDING', (0,0), (-1,-1), 5),
                ('BOTTOMPADDING', (0,0), (-1,-1), 5),
            ]))
            elements.append(t_sum)

            doc.build(elements)

            st.download_button(
                label="📥 見積書PDFをダウンロード",
                data=pdf_buffer.getvalue(),
                file_name=f"見積書_{file_timestamp}.pdf",
                mime="application/pdf"
            )
        except Exception as e:
            st.error(f"PDF生成エラー: {e}")

    # 履歴からの注文請書発行セクション
    st.write("---")
    st.subheader("📜 過去の見積・受発注履歴一覧")
    if os.path.exists(HISTORY_FILE):
        hist_df = pd.read_csv(HISTORY_FILE)
        if not hist_df.empty:
            st.dataframe(hist_df)
        else:
            st.info("履歴データはまだありません。")
    else:
        st.info("履歴ファイルがまだ存在しません。")

# ==========================================
# 画面 4：粗利計算・利益管理
# ==========================================
elif app_mode == "📊 4. 粗利計算・利益管理":
    st.title("📊 粗利計算・利益管理")
    if os.path.exists(HISTORY_FILE):
        hist_df = pd.read_csv(HISTORY_FILE)
        if not hist_df.empty and "売上合計(税抜)" in hist_df.columns:
            total_sales = hist_df["売上合計(税抜)"].sum()
            total_cost_body = hist_df["自動ボディ原価"].sum() if "自動ボディ原価" in hist_df.columns else 0
            total_cost_process = hist_df["自動加工費"].sum() if "自動加工費" in hist_df.columns else 0
            total_cost = total_cost_body + total_cost_process
            total_profit = total_sales - total_cost
            overall_margin = (total_profit / total_sales * 100) if total_sales > 0 else 0

            col_p1, col_p2, col_p3 = st.columns(3)
            col_p1.metric("総売上 (税抜)", f"¥ {total_sales:,}")
            col_p2.metric("総原価 (ボディ+加工)", f"¥ {total_cost:,}")
            col_p3.metric("総粗利益", f"¥ {total_profit:,} ({overall_margin:.1f}%)")

            st.write("---")
            st.subheader("案件別利益一覧")
            st.dataframe(hist_df)
        else:
            st.info("集計可能なデータがありません。")
    else:
        st.info("履歴データがありません。")

# ==========================================
# 画面 5：請求書作成・発行
# ==========================================
elif app_mode == "📄 5. 請求書作成・発行":
    st.title("📄 請求書作成・発行（インボイス対応）")
    if os.path.exists(HISTORY_FILE):
        hist_df = pd.read_csv(HISTORY_FILE)
        if not hist_df.empty:
            selected_row_idx = st.selectbox("請求書を発行する案件を選択", range(len(hist_df)), format_func=lambda x: f"{hist_df.loc[x, '日時']} - {hist_df.loc[x, '案件名']} ({hist_df.loc[x, '顧客名']})")
            
            target_row = hist_df.loc[selected_row_idx]
            st.write(f"**選択中案件:** {target_row['案件名']}")
            st.write(f"**お取引先:** {target_row['顧客名']}")
            st.write(f"**売上金額(税抜):** ¥ {target_row['売上合計(税抜)']:,}")

            if st.button("📄 インボイス対応請求書PDFを作成"):
                st.success("（請求書発行のロジックがここに適用されます）")
        else:
            st.info("請求書を発行できる履歴がありません。")
    else:
        st.info("履歴ファイルが存在しません。")