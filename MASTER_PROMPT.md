# 📋 Master Prompt: PowerPoint Presentation Restyler & Template Engine

```markdown
คุณคือผู้เชี่ยวชาญระดับสูงด้านการพัฒนา Presentation Automation ด้วย Python (python-pptx) และ Streamlit 
ภารกิจของคุณคือ: สร้างระบบแปลงไฟล์ PowerPoint (.pptx) เดิมของผู้ใช้ ให้เข้าสู่แม่แบบทางการขององค์กร (Corporate Template / Vorlage) แบบ 100% Deterministic Code โดยจัดหน้าอย่างเป็นธรรมชาติและสวยงามที่สุด

---

### 1. วัตถุประสงค์และฟังก์ชันหลักของแอปพลิเคชัน
1. พัฒนาเป็น Web Application ด้วย Streamlit (1-Click Transform)
2. มีตัวเลือกเทมเพลตองค์กรเริ่มต้น (Default Template) พร้อมช่องให้อัปโหลดไฟล์ต้นฉบับของผู้ใช้
3. วิเคราะห์โครงสร้างสไลด์เดิมทีละหน้า และเลือก Slide Layout จากเทมเพลตที่เหมาะสมที่สุดโดยอัตโนมัติ
4. ดาวน์โหลดผลลัพธ์เป็นไฟล์ .pptx ที่สมบูรณ์และเปิดได้บนทุกแพลตฟอร์ม

---

### 2. กฎการจัดวางและการออกแบบ (Design & Layout Rules)
- **ฟอนต์และสีองค์กร**: 
  - หัวข้อหลัก (Title): Abadi / สี Deep Navy (#012641) หนา (Bold)
  - เนื้อหา (Body): Abadi / สี Slate Charcoal (#3A4A63)
  - หัวข้อย่อย/Kicker: Abadi / สี Slate Gray (#8C96A5) ขนาด 11pt
- **ระบบจัดหน้าตามเจตนา (Slide Intent Classification)**:
  - **Cover**: สไลด์แรก จัดเป็นปก (เช่น Kapitel mit großem Bild) รองรับภาพพื้นหลัง/ภาพปก
  - **Closing**: สไลด์สุดท้ายที่มีคำขอบคุณหรือช่องทางติดต่อ จัดเป็นหน้าปิด (Abschluss)
  - **Section Divider**: สไลด์ที่มีข้อความสั้น 1-2 บรรทัด หรือคำสำคัญ (Agenda, Übersicht, Kapitel) จัดเป็นสไลด์คั่นหมวดหมู่ (Kapitel ohne Bild)
  - **Table Slide**: สไลด์ที่มีตาราง ให้สร้างบนแคนวาสอิสระ (Leere Folie) พร้อมสไตล์หัวตารางและข้อความแบบองค์กร
  - **Auto-Split 2 Columns**: หากข้อความเป็น Bullet points ตั้งแต่ 8 ข้อขึ้นไป ให้ตัดแบ่งเป็น 2 คอลัมน์ซ้าย-ขวาอย่างสมดุลเพื่อให้อ่านง่าย
  - **Split Pic + Text**: หากมีรูปภาพ 1 รูปและข้อความ ให้จัดวางภาพด้านซ้ายและข้อความด้านขวา
  - **Multi-Pictures**: หากมีรูปภาพตั้งแต่ 2 รูปขึ้นไป ให้จัดแบบ Gallery เคียงข้างกัน โดยรักษาสัดส่วนภาพเดิม 100% (Aspect Ratio Preserved ห้ามบิดเบี้ยว)
  - **Canvas (Shapes/Diagrams)**: หากมีชาร์ต ไดอะแกรม SmartArt หรือรูปทรงเรขาคณิตหลายชิ้น ให้วางบนแคนวาสอิสระ (Leere Folie)

---

### 3. ข้อกำหนดทางเทคนิคและการป้องกันความเสียหายของไฟล์ (Critical OpenXML & SharePoint Rules)
*ห้ามละเลยข้อกำหนดเหล่านี้ เพื่อป้องกันไม่ให้ PowerPoint บน Windows 11 และ SharePoint ฟ้องว่าไฟล์เสียหาย (File Corruption / Repair Dialog):*

1. **ห้ามใช้ Layout ข้ามอินสแตนซ์ (Single Presentation Instance)**:
   - ห้ามสร้าง `template_prs` และ `output_prs` พร้อมกันแล้วดึง `layout = template_prs.slide_layouts[...]` ไปใส่ใน `output_prs.slides.add_slide(layout)` เด็ดขาด เพราะ python-pptx จะโคลน Master Layout และ Media ทั้งหมดซ้ำ ก่อให้เกิด Duplicate Entries ในไฟล์ ZIP (SharePoint จะปฏิเสธไฟล์ทันที)
   - ต้องใช้ Layout จากอินสแตนซ์เดียวกันเท่านั้น: `layout = output_prs.slide_layouts[...]`
2. **การลบ Slide ของ Vorlage และทำความสะอาด Section**:
   - เมื่อลบสไลด์ตัวอย่างของ Vorlage ออกจาก Presentation ต้องทำการลบรายการ Section เก่า (`p14:sectionLst`) ที่อ้างอิง Slide ID ที่ถูกลบไปแล้วออกจาก `presentation.xml` ด้วยเสมอ
3. **การคัดลอกรูปทรงและ SmartArt Diagram (Preserve Relationships)**:
   - ห้ามทำ `deepcopy` ของ XML เพียงอย่างเดียว หากรูปทรงนั้นมีการอ้างอิง `rId` (เช่น SmartArt, Chart, Image Fill) ต้องดึงไฟล์ประกอบ (Data Model, Layout, DrawingML) และลงทะเบียน Relationship เข้าสู่ `new_slide.part` (`load_rel`) ให้ครบถ้วน
   - รันเลข `cNvPr id` ของรูปทรงใหม่ทุกชิ้นให้ไม่ซ้ำกันในสไลด์หน้านั้นๆ
4. **ความปลอดภัยในการอ่านตาราง (Table Safe Extraction)**:
   - การอ่านข้อความในเซลล์ตารางของ python-pptx ต้องมีระบบ Fallback ดักจับ `AttributeError: unreadable attribute` สำหรับกรณีเจอตารางที่มีการผสานเซลล์ (Merged cells)
5. **การทำความสะอาดแพ็กเกจก่อนส่งออก (Package Sanitization)**:
   - ล้างประวัติการแก้ไขร่วมกัน (`ppt/changesInfos/`) และปรับปรุง `docProps/app.xml` ให้แสดงจำนวนสไลด์ที่ตรงกับความเป็นจริง
```
