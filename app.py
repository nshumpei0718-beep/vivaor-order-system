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
EXCEL_FILE = "メーカー別品番価格リスト.xlsx"
COMPANY_PRESET_FILE = "company_presets.csv"
BANK_PRESET_FILE = "bank_presets.csv"

# 日本語フォントの設定（クロスプラットフォーム対応・安全なフォールバック）
FONT_NAME = 'Helvetica'
try:
    # Windows環境の場合のパスチェック
    if os.name == 'nt':
        if os.path.exists("C:/Windows/Fonts/meiryo.ttc"):
            pdfmetrics.registerFont(TTFont('Meiryo', 'C:/Windows/Fonts/meiryo.ttc'))
            FONT_NAME = 'Meiryo'
        elif os.path.exists("C:/Windows/Fonts/msgothic.ttc"):
            pdfmetrics.registerFont(TTFont('Msgothic', 'C:/Windows/Fonts/msgothic.ttc'))
            FONT_NAME = 'Msgothic'
    else:
        # Linux（Streamlit Cloud）等の場合、OS標準の日本語フォントがあれば登録を試みる
        for font_path in ["/usr/share/fonts/truetype/fonts-japanese-gothic.ttf", "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc"]:
            if os.path.exists(font_path):
                pdfmetrics.registerFont(TTFont('JapaneseFont', font_path))
                FONT_NAME = 'JapaneseFont'
                break
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
                ('TOPPADDING', (0,0), (-1,-1), 4),
                ('BOTTOMPADDING', (0,0), (-1,-1), 4),
            ]))
            elements.append(t_sum)
            elements.append(Spacer(1, 10))

            remark_data = [
                [Paragraph("<b>備考欄：</b>", meta_s)],
                [Paragraph("", meta_s)],
                [Paragraph("", meta_s)],
                [Paragraph("", meta_s)]
            ]
            t_remark = Table(remark_data, colWidths=[300], rowHeights=[15, 18, 18, 18])
            t_remark.setStyle(TableStyle([
                ('BOX', (0,0), (-1,-1), 0.5, colors.HexColor('#2E6EA5')),
                ('VALIGN', (0,0), (-1,-1), 'TOP'),
                ('TOPPADDING', (0,0), (-1,-1), 4),
            ]))
            elements.append(t_remark)

            doc.build(elements)

            st.download_button(
                label="📥 見積書PDFを保存",
                data=pdf_buffer.getvalue(),
                file_name=f"お見積書_{file_timestamp}.pdf",
                mime="application/pdf"
            )
        except Exception as e:
            st.error(f"見積書生成エラー: {e}")

    # ==========================================
    # 📚 過去の販売・見積履歴 ＆ 編集・削除・請書発行機能
    # ==========================================
    st.write("---")
    st.subheader("📚 過去の販売・見積履歴（編集・削除・請書発行が可能）")

    with st.expander("⚙️ 【事前登録】自社情報・振込先情報のマスタ管理（追加・編集）"):
        p_tab1, p_tab2 = st.tabs(["🏢 自社情報プリセット", "🏦 振込先情報プリセット"])
        
        with p_tab1:
            df_cp = pd.read_csv(COMPANY_PRESET_FILE)
            st.dataframe(df_cp, use_container_width=True)
            with st.form("add_company_preset_form"):
                st.markdown("**新規自社情報の追加**")
                new_cp_name = st.text_input("プリセット名（例: 本社、東京オフィス）")
                new_cp_cname = st.text_input("自社名", value="〇〇株式会社")
                new_cp_caddr = st.text_input("自社住所", value="〒123-4567 〇〇県〇〇市...")
                new_cp_cstaff = st.text_input("担当者名", value="担当 太郎")
                submitted_cp = st.form_submit_button("➕ この自社情報をプリセットに登録")
                if submitted_cp and new_cp_name:
                    new_row_df = pd.DataFrame([{"名称": new_cp_name, "自社名": new_cp_cname, "自社住所": new_cp_caddr, "自社担当": new_cp_cstaff}])
                    df_cp = pd.concat([df_cp, new_row_df], ignore_index=True)
                    df_cp.to_csv(COMPANY_PRESET_FILE, index=False, encoding="utf-8-sig")
                    st.success("自社情報を登録しました！")
                    st.rerun()

        with p_tab2:
            df_bp = pd.read_csv(BANK_PRESET_FILE)
            st.dataframe(df_bp, use_container_width=True)
            with st.form("add_bank_preset_form"):
                st.markdown("**新規振込先情報の追加**")
                new_bp_name = st.text_input("プリセット名（例: メイン口座、〇〇銀行）")
                new_bp_bank = st.text_input("銀行名", value="〇〇銀行")
                new_bp_branch = st.text_input("支店名", value="〇〇支店")
                new_bp_acc = st.text_input("口座番号", value="普通 1234567")
                new_bp_holder = st.text_input("口座名義", value="カ）〇〇カンパニー")
                submitted_bp = st.form_submit_button("➕ この振込先をプリセットに登録")
                if submitted_bp and new_bp_name:
                    new_row_df = pd.DataFrame([{"名称": new_bp_name, "銀行名": new_bp_bank, "支店名": new_bp_branch, "口座番号": new_bp_acc, "口座名義": new_bp_holder}])
                    df_bp = pd.concat([df_bp, new_row_df], ignore_index=True)
                    df_bp.to_csv(BANK_PRESET_FILE, index=False, encoding="utf-8-sig")
                    st.success("振込先情報を登録しました！")
                    st.rerun()

    if os.path.exists(HISTORY_FILE):
        history_df = pd.read_csv(HISTORY_FILE)
        for col_n, default_v in [("自社名", "〇〇株式会社"), ("自社住所", "〒123-4567 〇〇県〇〇市..."), ("自社担当", "担当 太郎"), ("銀行名", "〇〇銀行"), ("支店名", "〇〇支店"), ("口座番号", "普通 1234567"), ("口座名義", "カ）〇〇カンパニー")]:
            if col_n not in history_df.columns:
                history_df[col_n] = default_v

        if not history_df.empty:
            df_cp_load = pd.read_csv(COMPANY_PRESET_FILE)
            df_bp_load = pd.read_csv(BANK_PRESET_FILE)

            for h_idx, h_row in history_df.iterrows():
                with st.expander(f"📁 [{h_row['日時']}] 案件名: {h_row['案件名']} (顧客: {h_row['顧客名']} / 金額: {int(h_row['売上合計(税抜)']):,}円)"):
                    
                    st.markdown("#### ✏️ この履歴の編集・削除・自社/振込先情報の変更")
                    
                    edit_col1, edit_col2 = st.columns(2)
                    with edit_col1:
                        edit_proj_name = st.text_input("案件名編集", value=str(h_row['案件名']), key=f"edit_proj_{h_idx}")
                        edit_cust_name = st.text_input("顧客名編集", value=str(h_row['顧客名']), key=f"edit_cust_{h_idx}")
                        edit_print_m = st.text_input("プリント方式編集", value=str(h_row['プリント方式']), key=f"edit_print_{h_idx}")
                        
                        st.markdown("##### 🏢 自社情報（請書右上）")
                        comp_mode = st.radio(f"自社情報の入力方法 #{h_idx}", ["手入力", "事前登録から選択"], key=f"comp_mode_{h_idx}", horizontal=True)
                        
                        if comp_mode == "事前登録から選択" and not df_cp_load.empty:
                            c_options = df_cp_load["名称"].tolist()
                            selected_c_preset = st.selectbox(f"登録済み自社情報 #{h_idx}", c_options, key=f"sel_c_pre_{h_idx}")
                            matched_cp_row = df_cp_load[df_cp_load["名称"] == selected_c_preset].iloc[0]
                            default_cn = matched_cp_row["自社名"]
                            default_ca = matched_cp_row["自社住所"]
                            default_cs = matched_cp_row["自社担当"]
                        else:
                            default_cn = str(h_row.get('自社名', '〇〇株式会社'))
                            default_ca = str(h_row.get('自社住所', '〒123-4567 〇〇県〇〇市...'))
                            default_cs = str(h_row.get('自社担当', '担当 太郎'))

                        edit_c_name = st.text_input("自社名", value=default_cn, key=f"edit_c_name_{h_idx}")
                        edit_c_addr = st.text_input("自社住所", value=default_ca, key=f"edit_c_addr_{h_idx}")
                        edit_c_staff = st.text_input("担当者名", value=default_cs, key=f"edit_c_staff_{h_idx}")

                    with edit_col2:
                        edit_qty = st.number_input("数量編集", min_value=0, value=int(h_row['数量']), key=f"edit_qty_{h_idx}")
                        edit_sales = st.number_input("売上合計(税抜)編集", min_value=0, value=int(h_row['売上合計(税抜)']), step=100, key=f"edit_sales_{h_idx}")
                        
                        st.markdown("##### 🏦 振込先情報（請書右上）")
                        bank_mode = st.radio(f"振込先の入力方法 #{h_idx}", ["手入力", "事前登録から選択"], key=f"bank_mode_{h_idx}", horizontal=True)
                        
                        if bank_mode == "事前登録から選択" and not df_bp_load.empty:
                            b_options = df_bp_load["名称"].tolist()
                            selected_b_preset = st.selectbox(f"登録済み振込先 #{h_idx}", b_options, key=f"sel_b_pre_{h_idx}")
                            matched_bp_row = df_bp_load[df_bp_load["名称"] == selected_b_preset].iloc[0]
                            default_bn = matched_bp_row["銀行名"]
                            default_bb = matched_bp_row["支店名"]
                            default_ba = matched_bp_row["口座番号"]
                            default_bh = matched_bp_row["口座名義"]
                        else:
                            default_bn = str(h_row.get('銀行名', '〇〇銀行'))
                            default_bb = str(h_row.get('支店名', '〇〇支店'))
                            default_ba = str(h_row.get('口座番号', '普通 1234567'))
                            default_bh = str(h_row.get('口座名義', 'カ）〇〇カンパニー'))

                        edit_bank = st.text_input("銀行名", value=default_bn, key=f"edit_bank_{h_idx}")
                        edit_branch = st.text_input("支店名", value=default_bb, key=f"edit_branch_{h_idx}")
                        edit_account = st.text_input("口座番号", value=default_ba, key=f"edit_account_{h_idx}")
                        edit_holder = st.text_input("口座名義", value=default_bh, key=f"edit_holder_{h_idx}")

                    btn_col_e1, btn_col_e2 = st.columns(2)
                    with btn_col_e1:
                        if st.button("💾 変更を保存する", key=f"save_hist_{h_idx}"):
                            history_df.at[h_idx, '案件名'] = edit_proj_name
                            history_df.at[h_idx, '顧客名'] = edit_cust_name
                            history_df.at[h_idx, 'プリント方式'] = edit_print_m
                            history_df.at[h_idx, '数量'] = edit_qty
                            history_df.at[h_idx, '売上合計(税抜)'] = edit_sales
                            history_df.at[h_idx, '自社名'] = edit_c_name
                            history_df.at[h_idx, '自社住所'] = edit_c_addr
                            history_df.at[h_idx, '自社担当'] = edit_c_staff
                            history_df.at[h_idx, '銀行名'] = edit_bank
                            history_df.at[h_idx, '支店名'] = edit_branch
                            history_df.at[h_idx, '口座番号'] = edit_account
                            history_df.at[h_idx, '口座名義'] = edit_holder
                            
                            history_df.to_csv(HISTORY_FILE, index=False, encoding="utf-8-sig")
                            st.success("履歴および自社・振込先情報を更新しました！")
                            st.rerun()
                    with btn_col_e2:
                        if st.button("🗑️ この履歴を削除する", key=f"del_hist_{h_idx}"):
                            history_df = history_df.drop(h_idx).reset_index(drop=True)
                            history_df.to_csv(HISTORY_FILE, index=False, encoding="utf-8-sig")
                            st.success("履歴を削除しました！")
                            st.rerun()

                    st.write("---")
                    
                    col_h_info, col_h_btn = st.columns([3, 1])
                    with col_h_info:
                        st.markdown(
                            f"**顧客名:** {h_row['顧客名']}<br/>"
                            f"**プリント方式:** {h_row['プリント方式']} / **数量:** {int(h_row['数量'])}枚<br/>"
                            f"**売上合計(税抜):** {int(h_row['売上合計(税抜)']):,} 円",
                            unsafe_allow_html=True
                        )
                    with col_h_btn:
                        if st.button("📄 請書を発行する", key=f"issue_ord_from_hist_{h_idx}"):
                            st.session_state[f"gen_ord_{h_idx}"] = True

                    if st.session_state.get(f"gen_ord_{h_idx}", False):
                        st.success("✨ 請書（注文請書）のデータ準備ができました！")
                        try:
                            pdf_ord_buffer = io.BytesIO()
                            doc_o = SimpleDocTemplate(
                                pdf_ord_buffer, 
                                pagesize=portrait(A4), 
                                rightMargin=25, leftMargin=25, 
                                topMargin=25, bottomMargin=25
                            )
                            
                            t_style = ParagraphStyle('OrdTitle', fontName=FONT_NAME, fontSize=16, textColor=colors.white, alignment=0, spaceBefore=4, spaceAfter=4)
                            meta_s = ParagraphStyle('OrdMeta', fontName=FONT_NAME, fontSize=9)
                            right_meta_s = ParagraphStyle('OrdRMeta', fontName=FONT_NAME, fontSize=9, alignment=2)
                            cust_large_s = ParagraphStyle('OrdCustLarge', fontName=FONT_NAME, fontSize=13, leading=16)

                            cell_s = ParagraphStyle('OrdCell', fontName=FONT_NAME, fontSize=8)
                            cell_c = ParagraphStyle('OrdCellC', fontName=FONT_NAME, fontSize=8, alignment=1)
                            cell_r = ParagraphStyle('OrdCellR', fontName=FONT_NAME, fontSize=8, alignment=2)
                            hdr_s = ParagraphStyle('OrdHdr', fontName=FONT_NAME, fontSize=8, textColor=colors.white, alignment=1)
                            
                            bank_hdr_s = ParagraphStyle('BankHdr', fontName=FONT_NAME, fontSize=8, textColor=colors.white, alignment=1)
                            bank_val_s = ParagraphStyle('BankVal', fontName=FONT_NAME, fontSize=8, alignment=0)
                            
                            amt_title_s = ParagraphStyle('AmtT', fontName=FONT_NAME, fontSize=11, textColor=colors.white, alignment=1)
                            amt_val_s = ParagraphStyle('AmtV', fontName=FONT_NAME, fontSize=14, textColor=colors.black, alignment=2)

                            sum_hdr_style = ParagraphStyle('SumHdr', fontName=FONT_NAME, fontSize=8, textColor=colors.white, alignment=1)
                            sum_val_style = ParagraphStyle('SumVal', fontName=FONT_NAME, fontSize=8, textColor=colors.black, alignment=2)

                            elements_o = []

                            header_banner_data = [
                                [Paragraph("<b>注 文 請 書</b>", t_style), Paragraph(f"No : ORD-{h_idx}-{file_timestamp}<br/>発行日 : {datetime.now().strftime('%Y年%m月%d日')}", right_meta_s)]
                            ]
                            t_banner = Table(header_banner_data, colWidths=[300, 245])
                            t_banner.setStyle(TableStyle([
                                ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#2E6EA5')),
                                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                                ('BOTTOMPADDING', (0,0), (-1,-1), 8),
                                ('TOPPADDING', (0,0), (-1,-1), 8),
                            ]))
                            elements_o.append(t_banner)
                            elements_o.append(Spacer(1, 10))

                            cust_text = f"<b>{h_row['顧客名']}</b>"
                            
                            c_name_val = str(h_row.get('自社名', '〇〇株式会社'))
                            c_addr_val = str(h_row.get('自社住所', '〒123-4567 〇〇県〇〇市...'))
                            c_staff_val = str(h_row.get('自社担当', '担当 太郎'))

                            b_name_val = str(h_row.get('銀行名', '〇〇銀行'))
                            b_branch_val = str(h_row.get('支店名', '〇〇支店'))
                            b_acc_val = str(h_row.get('口座番号', '普通 1234567'))
                            b_holder_val = str(h_row.get('口座名義', 'カ）〇〇カンパニー'))

                            company_outside_info = [
                                Paragraph(f"<b>{c_name_val}</b>", meta_s),
                                Paragraph(c_addr_val, meta_s),
                                Paragraph(c_staff_val, meta_s),
                            ]

                            bank_info_data = [
                                [Paragraph("銀 行", bank_hdr_s), Paragraph(b_name_val, bank_val_s)],
                                [Paragraph("支 店", bank_hdr_s), Paragraph(b_branch_val, bank_val_s)],
                                [Paragraph("口座番号", bank_hdr_s), Paragraph(b_acc_val, bank_val_s)],
                                [Paragraph("口座名義", bank_hdr_s), Paragraph(b_holder_val, bank_val_s)],
                            ]
                            t_bank = Table(bank_info_data, colWidths=[55, 150])
                            t_bank.setStyle(TableStyle([
                                ('BACKGROUND', (0,0), (0,-1), colors.HexColor('#2E6EA5')),
                                ('TEXTCOLOR', (0,0), (0,-1), colors.white),
                                ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#2E6EA5')),
                                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                                ('TOPPADDING', (0,0), (-1,-1), 2),
                                ('BOTTOMPADDING', (0,0), (-1,-1), 2),
                            ]))

                            right_column_content = [
                                company_outside_info[0],
                                company_outside_info[1],
                                company_outside_info[2],
                                Spacer(1, 6),
                                t_bank
                            ]

                            top_info_data = [
                                [Paragraph(cust_text, cust_large_s), right_column_content]
                            ]
                            t_top = Table(top_info_data, colWidths=[310, 235])
                            t_top.setStyle(TableStyle([
                                ('VALIGN', (0,0), (-1,-1), 'TOP'),
                            ]))
                            elements_o.append(t_top)
                            elements_o.append(Spacer(1, 10))

                            elements_o.append(Paragraph(f"<b>件名：{h_row['案件名']}</b>", meta_s))
                            elements_o.append(Spacer(1, 4))
                            elements_o.append(Paragraph("下記の通り、ご注文を承りました。", meta_s))
                            elements_o.append(Spacer(1, 8))

                            net_amount = int(h_row['売上合計(税抜)'])
                            tax_amount = int(net_amount * 0.1)
                            total_inc_tax = net_amount + tax_amount

                            amt_box_data = [
                                [Paragraph("ご請求金額", amt_title_s), Paragraph(f"¥ {total_inc_tax:,} -", amt_val_s)]
                            ]
                            t_amt = Table(amt_box_data, colWidths=[120, 220])
                            t_amt.setStyle(TableStyle([
                                ('BACKGROUND', (0,0), (0,0), colors.HexColor('#2E6EA5')),
                                ('BOX', (0,0), (-1,-1), 1, colors.HexColor('#2E6EA5')),
                                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                                ('TOPPADDING', (0,0), (-1,-1), 6),
                                ('BOTTOMPADDING', (0,0), (-1,-1), 6),
                            ]))

                            notice_p = Paragraph("<font size=7>※お振込手数料は御社ご負担にてお願いします。</font>", meta_s)
                            
                            t_amt_layout = Table([[t_amt, notice_p]], colWidths=[350, 195])
                            t_amt_layout.setStyle(TableStyle([('VALIGN', (0,0), (-1,-1), 'BOTTOM')]))
                            elements_o.append(t_amt_layout)
                            elements_o.append(Spacer(1, 12))

                            h_qty = int(h_row['数量'])
                            h_unit_p = int(net_amount / h_qty) if h_qty > 0 else net_amount

                            table_data = [
                                [Paragraph("商品名 ／ 品名", hdr_s), Paragraph("単価", hdr_s), Paragraph("数量", hdr_s), Paragraph("金額", hdr_s)]
                            ]
                            table_data.append([
                                Paragraph(f"<b>{h_row['案件名']}</b>（{h_row['プリント方式']}）", cell_s),
                                Paragraph(f"{h_unit_p:,} 円", cell_r),
                                Paragraph(str(h_qty), cell_c),
                                Paragraph(f"{net_amount:,} 円", cell_r)
                            ])
                            for _ in range(9):
                                table_data.append([Paragraph("", cell_s), Paragraph("", cell_r), Paragraph("", cell_c), Paragraph("", cell_r)])

                            t_details = Table(table_data, colWidths=[264, 80, 60, 141], rowHeights=[20] + [18]*10)
                            
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
                            elements_o.append(t_details)
                            elements_o.append(Spacer(1, 8))

                            sum_table_data = [
                                [
                                    Paragraph("小計 (税抜)", sum_hdr_style), Paragraph(f"¥ {net_amount:,}", sum_val_style),
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
                                ('TOPPADDING', (0,0), (-1,-1), 4),
                                ('BOTTOMPADDING', (0,0), (-1,-1), 4),
                            ]))
                            elements_o.append(t_sum)
                            elements_o.append(Spacer(1, 10))

                            remark_data = [
                                [Paragraph("<b>備考欄：</b>", meta_s)],
                                [Paragraph("", meta_s)],
                                [Paragraph("", meta_s)],
                                [Paragraph("", meta_s)]
                            ]
                            t_remark = Table(remark_data, colWidths=[300], rowHeights=[15, 18, 18, 18])
                            t_remark.setStyle(TableStyle([
                                ('BOX', (0,0), (-1,-1), 0.5, colors.HexColor('#2E6EA5')),
                                ('VALIGN', (0,0), (-1,-1), 'TOP'),
                                ('TOPPADDING', (0,0), (-1,-1), 4),
                            ]))
                            elements_o.append(t_remark)

                            doc_o.build(elements_o)

                            st.download_button(
                                label="📥 【画像イメージ準拠の請書PDF】をダウンロード",
                                data=pdf_ord_buffer.getvalue(),
                                file_name=f"注文請書_{h_row['案件名']}.pdf",
                                mime="application/pdf",
                                key=f"dl_ord_hist_btn_{h_idx}"
                            )
                        except Exception as ex:
                            st.error(f"注文請書PDF生成エラー: {ex}")
        else:
            st.info("履歴データは空です。")
    else:
        st.info("過去の履歴データはまだありません。")

# ==========================================
# 画面 4：粗利計算・利益管理（全案件一覧対応版）
# ==========================================
elif app_mode == "📊 4. 粗利計算・利益管理":
    st.title("📊 案件別 粗利計算・利益管理")
    if not os.path.exists(HISTORY_FILE):
        st.info("⚠️ 過去の履歴データがありません。")
    else:
        history_df = pd.read_csv(HISTORY_FILE)
        
        # 全案件の合算・一覧表示セクション
        st.subheader("📈 全案件の利益サマリー・一覧")
        
        summary_rows = []
        total_all_sales = 0
        total_all_cost = 0
        total_all_profit = 0

        for idx, row in history_df.iterrows():
            s_val = float(row.get("売上合計(税抜)", 0))
            b_cost = float(row.get("自動ボディ原価", 0))
            p_cost = float(row.get("自動加工費", 0))
            c_val = b_cost + p_cost
            p_val = s_val - c_val
            m_val = (p_val / s_val * 100) if s_val > 0 else 0

            total_all_sales += s_val
            total_all_cost += c_val
            total_all_profit += p_val

            summary_rows.append({
                "日時": row.get("日時", ""),
                "案件名": row.get("案件名", ""),
                "顧客名": row.get("顧客名", ""),
                "売上(税抜)": s_val,
                "原価合計": c_val,
                "粗利額": p_val,
                "利益率": f"{m_val:.1f}%"
            })

        df_summary = pd.DataFrame(summary_rows)

        total_all_margin = (total_all_profit / total_all_sales * 100) if total_all_sales > 0 else 0
        
        m_col1, m_col2, m_col3 = st.columns(3)
        with m_col1:
            st.metric("総売上 (税抜)", f"{total_all_sales:,.0f} 円")
        with m_col2:
            st.metric("総原価", f"{total_all_cost:,.0f} 円")
        with m_col3:
            st.metric("総粗利額 / 平均利益率", f"{total_all_profit:,.0f} 円", delta=f"{total_all_margin:.1f}%")

        st.markdown("##### 📋 全案件の明細一覧")
        st.dataframe(df_summary, use_container_width=True)

        st.write("---")
        st.subheader("🔍 個別案件の詳細・シミュレーション")
        
        project_options = [f"{row['日時']} - {row['案件名']} ({row['顧客名']})" for idx, row in history_df.iterrows()]
        selected_proj_idx = st.selectbox("対象案件を選択して個別調整", range(len(project_options)), format_func=lambda x: project_options[x])
        target_row = history_df.iloc[selected_proj_idx]
        
        col_c1, col_c2 = st.columns(2)
        with col_c1:
            body_total_cost = st.number_input("① ボディ仕入原価 総額 (税抜)", min_value=0, value=int(target_row["自動ボディ原価"]), step=100)
        with col_c2:
            print_process_cost = st.number_input("② プリント加工費 総額 (税抜)", min_value=0, value=int(target_row["自動加工費"]), step=100)

        total_sales = float(target_row["売上合計(税抜)"])
        total_cost = body_total_cost + print_process_cost
        gross_profit = total_sales - total_cost
        gross_profit_margin = (gross_profit / total_sales * 100) if total_sales > 0 else 0

        st.metric("選択案件の粗利額 (利益額)", f"{gross_profit:,.0f} 円", delta=f"{gross_profit_margin:.1f}% (利益率)")

# ==========================================
# 画面 5：請求書作成・発行
# ==========================================
elif app_mode == "📄 5. 請求書作成・発行":
    st.title("📄 請求書作成・発行（インボイス制度対応）")
    if not os.path.exists(HISTORY_FILE):
        st.info("⚠️ 過去の販売履歴データがありません。")
    else:
        history_df = pd.read_csv(HISTORY_FILE)
        invoice_options = [f"{row['日時']} - {row['案件名']} ({row['顧客名']})" for idx, row in history_df.iterrows()]
        selected_inv_idx = st.selectbox("請求書を発行する案件を選択", range(len(invoice_options)), format_func=lambda x: invoice_options[x])
        inv_row = history_df.iloc[selected_inv_idx]
        
        inv_project = st.text_input("請求書件名", value=inv_row["案件名"])
        inv_customer = st.text_input("請求先（顧客名）", value=inv_row["顧客名"])
        inv_amount = st.number_input("請求金額（税抜）", min_value=0, value=int(inv_row["売上合計(税抜)"]), step=100)
        
        if st.button("📥 請求書PDFをダウンロード"):
            try:
                pdf_inv_buffer = io.BytesIO()
                doc_i = SimpleDocTemplate(
                    pdf_inv_buffer, 
                    pagesize=portrait(A4), 
                    rightMargin=25, leftMargin=25, 
                    topMargin=25, bottomMargin=25
                )
                
                t_style = ParagraphStyle('InvTitle', fontName=FONT_NAME, fontSize=16, textColor=colors.white, alignment=0, spaceBefore=4, spaceAfter=4)
                meta_s = ParagraphStyle('InvMeta', fontName=FONT_NAME, fontSize=9)
                right_meta_s = ParagraphStyle('InvRMeta', fontName=FONT_NAME, fontSize=9, alignment=2)
                cust_large_s = ParagraphStyle('InvCustLarge', fontName=FONT_NAME, fontSize=13, leading=16)

                cell_s = ParagraphStyle('InvCell', fontName=FONT_NAME, fontSize=8)
                cell_c = ParagraphStyle('InvCellC', fontName=FONT_NAME, fontSize=8, alignment=1)
                cell_r = ParagraphStyle('InvCellR', fontName=FONT_NAME, fontSize=8, alignment=2)
                hdr_s = ParagraphStyle('InvHdr', fontName=FONT_NAME, fontSize=8, textColor=colors.white, alignment=1)
                
                bank_hdr_s = ParagraphStyle('BankHdr', fontName=FONT_NAME, fontSize=8, textColor=colors.white, alignment=1)
                bank_val_s = ParagraphStyle('BankVal', fontName=FONT_NAME, fontSize=8, alignment=0)
                
                amt_title_s = ParagraphStyle('AmtT', fontName=FONT_NAME, fontSize=11, textColor=colors.white, alignment=1)
                amt_val_s = ParagraphStyle('AmtV', fontName=FONT_NAME, fontSize=14, textColor=colors.black, alignment=2)

                sum_hdr_style = ParagraphStyle('SumHdr', fontName=FONT_NAME, fontSize=8, textColor=colors.white, alignment=1)
                sum_val_style = ParagraphStyle('SumVal', fontName=FONT_NAME, fontSize=8, textColor=colors.black, alignment=2)

                elements_i = []

                header_banner_data = [
                    [Paragraph("<b>御 請 求 書</b>", t_style), Paragraph(f"No : INV-{selected_inv_idx}-{datetime.now().strftime('%Y%m%d_%H%M%S')}<br/>発行日 : {datetime.now().strftime('%Y年%m月%d日')}", right_meta_s)]
                ]
                t_banner = Table(header_banner_data, colWidths=[300, 245])
                t_banner.setStyle(TableStyle([
                    ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#2E6EA5')),
                    ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                    ('BOTTOMPADDING', (0,0), (-1,-1), 8),
                    ('TOPPADDING', (0,0), (-1,-1), 8),
                ]))
                elements_i.append(t_banner)
                elements_i.append(Spacer(1, 10))

                cust_text = f"<b>{inv_customer}</b>"
                
                c_name_val = str(inv_row.get('自社名', '〇〇株式会社'))
                c_addr_val = str(inv_row.get('自社住所', '〒123-4567 〇〇県〇〇市...'))
                c_staff_val = str(inv_row.get('自社担当', '担当 太郎'))

                b_name_val = str(inv_row.get('銀行名', '〇〇銀行'))
                b_branch_val = str(inv_row.get('支店名', '〇〇支店'))
                b_acc_val = str(inv_row.get('口座番号', '普通 1234567'))
                b_holder_val = str(inv_row.get('口座名義', 'カ）〇〇カンパニー'))

                company_outside_info = [
                    Paragraph(f"<b>{c_name_val}</b>", meta_s),
                    Paragraph(c_addr_val, meta_s),
                    Paragraph(c_staff_val, meta_s),
                ]

                bank_info_data = [
                    [Paragraph("銀 行", bank_hdr_s), Paragraph(b_name_val, bank_val_s)],
                    [Paragraph("支 店", bank_hdr_s), Paragraph(b_branch_val, bank_val_s)],
                    [Paragraph("口座番号", bank_hdr_s), Paragraph(b_acc_val, bank_val_s)],
                    [Paragraph("口座名義", bank_hdr_s), Paragraph(b_holder_val, bank_val_s)],
                ]
                t_bank = Table(bank_info_data, colWidths=[55, 150])
                t_bank.setStyle(TableStyle([
                    ('BACKGROUND', (0,0), (0,-1), colors.HexColor('#2E6EA5')),
                    ('TEXTCOLOR', (0,0), (0,-1), colors.white),
                    ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#2E6EA5')),
                    ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                    ('TOPPADDING', (0,0), (-1,-1), 2),
                    ('BOTTOMPADDING', (0,0), (-1,-1), 2),
                ]))

                right_column_content = [
                    company_outside_info[0],
                    company_outside_info[1],
                    company_outside_info[2],
                    Spacer(1, 6),
                    t_bank
                ]

                top_info_data = [
                    [Paragraph(cust_text, cust_large_s), right_column_content]
                ]
                t_top = Table(top_info_data, colWidths=[310, 235])
                t_top.setStyle(TableStyle([
                    ('VALIGN', (0,0), (-1,-1), 'TOP'),
                ]))
                elements_i.append(t_top)
                elements_i.append(Spacer(1, 10))

                elements_i.append(Paragraph(f"<b>件名：{inv_project}</b>", meta_s))
                elements_i.append(Spacer(1, 4))
                elements_i.append(Paragraph("下記の通り、ご請求申し上げます。", meta_s))
                elements_i.append(Spacer(1, 8))

                net_amount = int(inv_amount)
                tax_amount = int(net_amount * 0.1)
                total_inc_tax = net_amount + tax_amount

                amt_box_data = [
                    [Paragraph("ご請求金額", amt_title_s), Paragraph(f"¥ {total_inc_tax:,} -", amt_val_s)]
                ]
                t_amt = Table(amt_box_data, colWidths=[120, 220])
                t_amt.setStyle(TableStyle([
                    ('BACKGROUND', (0,0), (0,0), colors.HexColor('#2E6EA5')),
                    ('BOX', (0,0), (-1,-1), 1, colors.HexColor('#2E6EA5')),
                    ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                    ('TOPPADDING', (0,0), (-1,-1), 6),
                    ('BOTTOMPADDING', (0,0), (-1,-1), 6),
                ]))

                notice_p = Paragraph("<font size=7>※お振込手数料は御社ご負担にてお願いします。</font>", meta_s)
                
                t_amt_layout = Table([[t_amt, notice_p]], colWidths=[350, 195])
                t_amt_layout.setStyle(TableStyle([('VALIGN', (0,0), (-1,-1), 'BOTTOM')]))
                elements_i.append(t_amt_layout)
                elements_i.append(Spacer(1, 12))

                i_qty = int(inv_row.get('数量', 1))
                i_unit_p = int(net_amount / i_qty) if i_qty > 0 else net_amount

                table_data = [
                    [Paragraph("商品名 ／ 品名", hdr_s), Paragraph("単価", hdr_s), Paragraph("数量", hdr_s), Paragraph("金額", hdr_s)]
                ]
                table_data.append([
                    Paragraph(f"<b>{inv_project}</b>（{inv_row.get('プリント方式', 'プリント')}）", cell_s),
                    Paragraph(f"{i_unit_p:,} 円", cell_r),
                    Paragraph(str(i_qty), cell_c),
                    Paragraph(f"{net_amount:,} 円", cell_r)
                ])
                for _ in range(9):
                    table_data.append([Paragraph("", cell_s), Paragraph("", cell_r), Paragraph("", cell_c), Paragraph("", cell_r)])

                t_details = Table(table_data, colWidths=[264, 80, 60, 141], rowHeights=[20] + [18]*10)
                
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
                elements_i.append(t_details)
                elements_i.append(Spacer(1, 8))

                sum_table_data = [
                    [
                        Paragraph("小計 (税抜)", sum_hdr_style), Paragraph(f"¥ {net_amount:,}", sum_val_style),
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
                    ('TOPPADDING', (0,0), (-1,-1), 4),
                    ('BOTTOMPADDING', (0,0), (-1,-1), 4),
                ]))
                elements_i.append(t_sum)
                elements_i.append(Spacer(1, 10))

                remark_data = [
                    [Paragraph("<b>備考欄：</b>", meta_s)],
                    [Paragraph("", meta_s)],
                    [Paragraph("", meta_s)],
                    [Paragraph("", meta_s)]
                ]
                t_remark = Table(remark_data, colWidths=[300], rowHeights=[15, 18, 18, 18])
                t_remark.setStyle(TableStyle([
                    ('BOX', (0,0), (-1,-1), 0.5, colors.HexColor('#2E6EA5')),
                    ('VALIGN', (0,0), (-1,-1), 'TOP'),
                    ('TOPPADDING', (0,0), (-1,-1), 4),
                ]))
                elements_i.append(t_remark)

                doc_i.build(elements_i)

                st.download_button(
                    label="📥 請求書PDFを保存",
                    data=pdf_inv_buffer.getvalue(),
                    file_name=f"御請求書_{inv_project}.pdf",
                    mime="application/pdf"
                )
            except Exception as e:
                st.error(f"請求書PDF生成エラー: {e}")