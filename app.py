import os
import re
import streamlit as st
from docx import Document
from docx.enum.section import WD_SECTION
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docxcompose.composer import Composer

# 強制設定節區頁碼從 1 開始獨立計算，且將 NUMPAGES 改為計算該節頁數 SECTIONPAGES
def setup_independent_section_page_numbering(doc):
    for section in doc.sections:
        sectPr = section._sectPr
        pgNumType = sectPr.find(qn('w:pgNumType'))
        if pgNumType is None:
            pgNumType = OxmlElement('w:pgNumType')
            sectPr.append(pgNumType)
        pgNumType.set(qn('w:start'), '1')
        
        # 替換頁尾中的總頁數欄位代碼：NUMPAGES -> SECTIONPAGES
        for footer in [section.footer, section.first_page_footer, section.even_page_footer]:
            if footer:
                for p in footer.paragraphs:
                    for fld in p._p.xpath('.//w:fldSimple'):
                        instr = fld.get(qn('w:instr'))
                        if instr and 'NUMPAGES' in instr:
                            fld.set(qn('w:instr'), instr.replace('NUMPAGES', 'SECTIONPAGES'))
                    for instrText in p._p.xpath('.//w:instrText'):
                        if instrText.text and 'NUMPAGES' in instrText.text:
                            instrText.text = instrText.text.replace('NUMPAGES', 'SECTIONPAGES')

# 具備 Regex 支援的精準文字替換
def replace_text_safely(paragraph, pattern_or_text, new_text, is_regex=False):
    if is_regex:
        if re.search(pattern_or_text, paragraph.text):
            full_text = re.sub(pattern_or_text, new_text, paragraph.text)
            _apply_text_to_paragraph(paragraph, full_text)
    else:
        if pattern_or_text in paragraph.text:
            full_text = paragraph.text.replace(pattern_or_text, new_text)
            _apply_text_to_paragraph(paragraph, full_text)

def _apply_text_to_paragraph(paragraph, full_text):
    if paragraph.runs:
        paragraph.runs[0].text = full_text
        for r in paragraph.runs[1:]:
            r.text = ""
    else:
        paragraph.text = full_text

def process_document_replacements(doc, replacements, regex_replacements=None):
    for p in doc.paragraphs:
        for old_txt, new_txt in replacements.items():
            replace_text_safely(p, old_txt, new_txt, is_regex=False)
        if regex_replacements:
            for pattern, new_txt in regex_replacements.items():
                replace_text_safely(p, pattern, new_txt, is_regex=True)
                
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for old_txt, new_txt in replacements.items():
                        replace_text_safely(p, old_txt, new_txt, is_regex=False)
                    if regex_replacements:
                        for pattern, new_txt in regex_replacements.items():
                            replace_text_safely(p, pattern, new_txt, is_regex=True)

st.set_page_config(page_title="H2U 永悅健康 - 全功能合約自動生成系統", layout="wide")
st.title("📄 H2U 全方位職場健康管理 - 全功能合約生成系統")

with st.form("contract_form"):
    st.subheader("一、客戶基本資訊（主約第1頁 & 末頁）")
    c1, c2, c3 = st.columns(3)
    with c1:
        client_name = st.text_input("甲方公司全名", "台灣積體電路製造股份有限公司")
        client_tax_id = st.text_input("統一編號", "22099118")
    with c2:
        client_rep = st.text_input("代表人/負責人", "魏哲家")
        client_address = st.text_input("公司地址", "新竹科學園區力行六路8號")
    with c3:
        sign_date = st.text_input("簽約日期", "2026年9月1日")
        industry_type = st.text_input("事業性質分類", "第一類高風險")

    st.subheader("二、合約期間與總價（主約第1~2頁）")
    c4, c5, c6 = st.columns(3)
    with c4:
        start_date_str = st.text_input("合作開始日期", "2026年9月1日")
        end_date_str = st.text_input("合作結束日期", "2027年8月31日")
    with c5:
        duration_months = st.text_input("合作總月份數", "12")
        total_amount = st.text_input("契約總價（新臺幣）", "1,000,000")
    with c6:
        payment_days = st.text_input("付款月結天數", "30")
        factory_name = st.text_input("廠區名稱", "新竹一廠")

    st.subheader("三、服務計費細節（主約第2頁 & 各附約）")
    c7, c8 = st.columns(2)
    with c7:
        employee_count = st.text_input("員工人數", "250")
        doctor_rate = st.text_input("醫師服務時薪（元/含稅）", "3,000")
        nurse_rate = st.text_input("護理人員服務時薪（元/含稅）", "1,000")
    with c8:
        doc_hours = st.text_input("醫師每月次數/小時", "2次/月、2小時/次")
        nurse_hours = st.text_input("護理師每月次數/小時", "4次/月、4小時/次")
        credit_limit = st.text_input("全景平台信用額度", "1,000,000")

    st.subheader("四、採購服務項目（自動勾選第1頁與末頁附件清單）")
    opt_sys = st.checkbox("H2U客戶健康管理系統使用授權")
    opt_health = st.checkbox("職場健康規劃服務")
    opt_exam = st.checkbox("健檢顧問服務")
    opt_pano = st.checkbox("H2U數位健康全景平台之使用授權")

    st.subheader("五、非標條款修改需求（若無可留空）")
    amendment_text = st.text_area("客戶要求修改之非開放條款內容（例如：保固期由1年改為2年）", "")

    submitted = st.form_submit_button("🚀 一鍵自動帶入全頁欄位並生成合約包")

if submitted:
    try:
        master_files = [f for f in os.listdir(".") if f.endswith(".docx") and "H2U全方位" in f and not f.startswith("~$") and not f.startswith("output_")]
        if not master_files:
            st.error("找不到主約 Word 檔案！")
            st.stop()

        doc_master = Document(master_files[0])

        master_replacements = {
            "請填入公司全名": client_name,
            "2025年9月1日": sign_date,
            "2026年8月31日": end_date_str,
            "00 個月": f"{duration_months} 個月",
            "OO廠": factory_name,
            "此後第幾類什麼型請自己填": industry_type,
            "月結【30】日內": f"月結【{payment_days}】日內",
            "本契約總價為新臺幣 (下同)      元整": f"本契約總價為新臺幣 (下同) {total_amount}元整",
            "甲    方：     ": f"甲    方：{client_name}",
            "代 表 人：     ": f"代 表 人：{client_rep}",
            "統一編號：     ": f"統一編號：{client_tax_id}",
            "地    址：     ": f"地    址：{client_address}",
        }

        # 針對可能包含不同空白字符的服務時薪進行正規表示式替換
        regex_replacements = {
            r"醫師實際服務時數乘以\s*元\(含稅\)": f"醫師實際服務時數乘以 {doctor_rate} 元(含稅)",
            r"護理人員實際服務時數乘以\s*元\(含稅\)": f"護理人員實際服務時數乘以 {nurse_rate} 元(含稅)",
        }

        # 方框打勾設定
        if opt_sys:
            master_replacements["☐ H2U客戶健康管理系統使用授權"] = "■ H2U客戶健康管理系統使用授權"
            master_replacements["☐  H2U客戶健康管理系統使用授權條款"] = "■  H2U客戶健康管理系統使用授權條款"
        if opt_health:
            master_replacements["☐ 職場健康規劃服務"] = "■ 職場健康規劃服務"
            master_replacements["☐  職場健康規劃服務條款"] = "■  職場健康規劃服務條款"
        if opt_exam:
            master_replacements["☐ 健檢顧問服務"] = "■ 健檢顧問服務"
            master_replacements["☐  健檢顧問服務條款"] = "■  健檢顧問服務條款"
        if opt_pano:
            master_replacements["☐ H2U數位健康全景平台之使用授權"] = "■ H2U數位健康全景平台之使用授權"
            master_replacements["☐  H2U數位健康全景平台之使用授權條款"] = "■  H2U數位健康全景平台之使用授權條款"

        process_document_replacements(doc_master, master_replacements, regex_replacements)
        setup_independent_section_page_numbering(doc_master)

        composer = Composer(doc_master)

        # 改用 list 結構，解決 dict key 覆蓋問題，確保所有勾選項均被讀取
        selected_attachments = []
        if opt_sys:
            selected_attachments.append("健康管理系統")
        if opt_health:
            selected_attachments.append("職場健康規劃")
        if opt_exam:
            selected_attachments.append("健檢顧問")
        if opt_pano:
            selected_attachments.append("全景平台")

        att_replacements = {
            "     ": client_name,
            "9,999,999": credit_limit,
            "1,000,000": credit_limit,
            "甲    方：     ": f"甲    方：{client_name}",
            "代 表 人：     ": f"代 表 人：{client_rep}",
            "統一編號：     ": f"統一編號：{client_tax_id}",
            "地    址：     ": f"地    址：{client_address}",
        }

        for keyword in selected_attachments:
            matched_files = [f for f in os.listdir(".") if keyword in f and f.endswith(".docx") and not f.startswith("~$")]
            if matched_files:
                att_doc = Document(matched_files[0])
                process_document_replacements(att_doc, att_replacements)
                setup_independent_section_page_numbering(att_doc)
                
                # 在合併前加入獨立新頁節區，避免表格與前文黏合
                composer.doc.add_section(WD_SECTION.NEW_PAGE)
                composer.append(att_doc)

        # 增補協議處理
        if amendment_text.strip():
            amend_files = [f for f in os.listdir(".") if "增補協議" in f and f.endswith(".docx") and not f.startswith("~$")]
            if amend_files:
                amend_doc = Document(amend_files[0])
                amend_replacements = {
                    "請填入原合約完整文件名稱": "H2U全方位職場健康管理服務契約",
                    "例如原合約第O條第X項第Y款改為：(如欲新增下一條請直接enter如擬於同條新增新項或換行請shift+enter)": f"原合約條款修訂如下：\n{amendment_text}",
                    "2025年7月10日": sign_date,
                    "2025年3月1日": sign_date,
                    "甲        方：     ": f"甲        方：{client_name}",
                    "代  表  人：     ": f"代  表  人：{client_rep}",
                    "統一編號：     ": f"統一編號：{client_tax_id}",
                    "地        址：     ": f"地        址：{client_address}",
                }
                process_document_replacements(amend_doc, amend_replacements)
                setup_independent_section_page_numbering(amend_doc)
                
                # 強制獨立分頁，隔開前份文件
                composer.doc.add_section(WD_SECTION.NEW_PAGE)
                composer.append(amend_doc)

        output_filename = f"{client_name}_完整合約包.docx"
        composer.save(output_filename)

        st.success("🎉 全頁欄位已成功帶入並生成完整合約包！")
        with open(output_filename, "rb") as file:
            st.download_button(
                label="📥 點擊下載完整 Word 合約包 (.docx)",
                data=file,
                file_name=output_filename,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

    except Exception as e:
        st.error(f"生成過程發生錯誤：{str(e)}")
