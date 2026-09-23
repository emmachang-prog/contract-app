import os
import streamlit as st
from docx import Document
from docxcompose.composer import Composer


# 支援跨 Run 完整字串精確替換，不破壞原有排版
def replace_text_safely(paragraph, old_text, new_text):
  if old_text not in paragraph.text:
    return

  # 如果整個段落剛好包含，且是在多個 run 中被拆散
  full_text = paragraph.text
  full_text = full_text.replace(old_text, new_text)

  # 記錄第一個 run 的樣式後重組文字
  if paragraph.runs:
    run0 = paragraph.runs[0]
    font_name = run0.font.name
    font_size = run0.font.size
    bold = run0.bold
    # 清空後續 run 的文字，把合併字串放進第一個 run
    run0.text = full_text
    for r in paragraph.runs[1:]:
      r.text = ""
  else:
    paragraph.text = full_text


def process_document_replacements(doc, replacements):
  for p in doc.paragraphs:
    for old_txt, new_txt in replacements.items():
      replace_text_safely(p, old_txt, new_txt)
  for table in doc.tables:
    for row in table.rows:
      for cell in row.cells:
        for p in cell.paragraphs:
          for old_txt, new_txt in replacements.items():
            replace_text_safely(p, old_txt, new_txt)


st.set_page_config(
    page_title="H2U 永悅健康 - 全功能合約自動生成系統", layout="wide"
)
st.title("📄 H2U 全方位職場健康管理 - 全功能合約生成系統")

with st.form("contract_form"):
  st.subheader("一、客戶基本資訊（主約第1頁 & 末頁）")
  c1, c2, c3 = st.columns(3)
  with c1:
    client_name = st.text_input("甲方公司全名", "台灣積體電路製造股份有限公司")
    client_tax_id = st.text_input("統一編號", "22099118")
  with c2:
    client_rep = st.text_input("代表人/負責人", "魏哲家")
    client_address = st.text_input(
        "公司地址", "新竹科學園區力行六路8號"
    )
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
  amendment_text = st.text_area(
      "客戶要求修改之非開放條款內容（例如：保固期由1年改為2年）", ""
  )

  submitted = st.form_submit_button("🚀 一鍵自動帶入全頁欄位並生成合約包")

if submitted:
  try:
    master_files = [
        f
        for f in os.listdir(".")
        if f.endswith(".docx")
        and "H2U全方位" in f
        and not f.startswith("~$")
        and not f.startswith("output_")
    ]
    if not master_files:
      st.error("找不到主約 Word 檔案！")
      st.stop()

    doc_master = Document(master_files[0])

    # 1. 主約替換字典：完全鎖定範本現有文字
    master_replacements = {
        "請填入公司全名": client_name,
        "2025年9月1日": sign_date,
        "2026年8月31日": end_date_str,
        "00 個月": f"{duration_months} 個月",
        "OO廠": factory_name,
        "新竹一廠": factory_name,
        "此後第幾類什麼型請自己填": industry_type,
        "月結【30】日內": f"月結【{payment_days}】日內",
        "甲    方：     ": f"甲    方：{client_name}",
        "代 表 人：     ": f"代 表 人：{client_rep}",
        "統一編號：     ": f"統一編號：{client_tax_id}",
        "地    址：     ": f"地    址：{client_address}",
    }

    # 勾選框替換
    if opt_sys:
      master_replacements[
          "☐ H2U客戶健康管理系統使用授權"
      ] = "■ H2U客戶健康管理系統使用授權"
      master_replacements[
          "☐  H2U客戶健康管理系統使用授權條款"
      ] = "■  H2U客戶健康管理系統使用授權條款"
    if opt_health:
      master_replacements["☐ 職場健康規劃服務"] = "■ 職場健康規劃服務"
      master_replacements["☐  職場健康規劃服務條款"] = "■  職場健康規劃服務條款"
    if opt_exam:
      master_replacements["☐ 健檢顧問服務"] = "■ 健檢顧問服務"
      master_replacements["☐  健檢顧問服務條款"] = "■  健檢顧問服務條款"
    if opt_pano:
      master_replacements[
          "☐ H2U數位健康全景平台之使用授權"
      ] = "■ H2U數位健康全景平台之使用授權"
      master_replacements[
          "☐  H2U數位健康全景平台之使用授權條款"
      ] = "■  H2U數位健康全景平台之使用授權條款"

    process_document_replacements(doc_master, master_replacements)

    composer = Composer(doc_master)

    # 2. 附約附加與文字替換
    attachment_map = {
        opt_sys: [
            f
            for f in os.listdir(".")
            if "健康管理系統" in f
            and f.endswith(".docx")
            and not f.startswith("~$")
        ],
        opt_health: [
            f
            for f in os.listdir(".")
            if "職場健康規劃" in f
            and f.endswith(".docx")
            and not f.startswith("~$")
        ],
        opt_exam: [
            f
            for f in os.listdir(".")
            if "健檢顧問" in f
            and f.endswith(".docx")
            and not f.startswith("~$")
        ],
        opt_pano: [
            f
            for f in os.listdir(".")
            if "全景平台" in f
            and f.endswith(".docx")
            and not f.startswith("~$")
        ],
    }

    att_replacements = {
        "     ": client_name,
        "9,999,999": credit_limit,
    }

    for is_selected, matched_files in attachment_map.items():
      if is_selected and matched_files:
        att_doc = Document(matched_files[0])
        process_document_replacements(att_doc, att_replacements)
        composer.append(att_doc)

    # 3. 增補協議處理
    if amendment_text.strip():
      amend_files = [
          f
          for f in os.listdir(".")
          if "增補協議" in f and f.endswith(".docx") and not f.startswith("~$")
      ]
      if amend_files:
        amend_doc = Document(amend_files[0])
        amend_replacements = {
            "請填入原合約完整文件名稱": (
                "H2U全方位職場健康管理服務契約"
            ),
            "例如原合約第O條第X項第Y款改為：(如欲新增下一條請直接enter如擬於同條新增新項或換行請shift+enter)": (
                f"原合約條款修訂如下：\n{amendment_text}"
            ),
            "2025年7月10日": sign_date,
            "2025年3月1日": sign_date,
            "甲        方：     ": f"甲        方：{client_name}",
            "代  表  人：     ": f"代  表  人：{client_rep}",
            "統一編號：     ": f"統一編號：{client_tax_id}",
            "地        址：     ": f"地        址：{client_address}",
        }
        process_document_replacements(amend_doc, amend_replacements)
        composer.append(amend_doc)

    output_filename = f"{client_name}_完整合約包.docx"
    composer.save(output_filename)

    st.success("🎉 全頁欄位已成功帶入並生成完整合約包！")
    with open(output_filename, "rb") as file:
      st.download_button(
          label="📥 點擊下載完整 Word 合約包 (.docx)",
          data=file,
          file_name=output_filename,
          mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      )

  except Exception as e:
    st.error(f"生成過程發生錯誤：{str(e)}")
