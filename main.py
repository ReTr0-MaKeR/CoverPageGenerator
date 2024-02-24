from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.label import Label
from kivy.uix.popup import Popup
from kivy.properties import StringProperty
from pathlib import Path
from docxtpl import DocxTemplate

class LayoutRoot(BoxLayout):
    ctitle_text = StringProperty("")  
    desg_text = StringProperty("")    

class LayoutApp(App):
    def build(self):
        return LayoutRoot()

    def on_start(self):
        self.course_mapping = {
            "CSE312": "Programming II",
            "EEE314": "Microprocessor & Interfacing Laboratory",
            "EEE316": "Power Syatem Analysis Laboratory",
            "EEE318": "Control System Laboratory",
            
        }

        self.teacher_mapping = {
            "Ms. Syeda Rukaiya Hossain": "Lecturer",
            "Ms. Tanjum Rahi Akanto": "Lecturer",
            "Mr. Md. Mahmudur Rahman": "Lecturer",
            "Dr. Srimanti Roychoudhury": "Associate Professor",
        }

        self.root.ids.ccode_spinner.values = list(self.course_mapping.keys())
        self.root.ids.tname_spinner.values = list(self.teacher_mapping.keys())

    def generate_document(self):
        document_path = Path(__file__).parent / "Cover_temp.docx"
        doc = DocxTemplate(document_path)

        
        no = self.root.ids.no_input.text
        expname = self.root.ids.expname_input.text
        ccode = self.root.ids.ccode_spinner.text
        ctitle = self.root.ctitle_text  
        tname = self.root.ids.tname_spinner.text
        desg = self.root.desg_text     
        name = self.root.ids.name_input.text
        sid = self.root.ids.sid_input.text
        batch = self.root.ids.batch_input.text
        lt = self.root.ids.lt_input.text
        section = self.root.ids.section_input.text
        expd = self.root.ids.expd_input.text
        subd = self.root.ids.subd_input.text

        context = {
            "No": no,
            "ExpName": expname,
            "CCode": ccode,
            "CTitle": ctitle,
            "TName": tname,
            "Designation": desg,
            "Name": name,
            "ID": sid,
            "Batch": batch,
            "LT": lt,
            "Section": section,
            "Exp_D": expd,
            "Sub_D": subd, 
        }

        
        doc.render(context)

        
        output_path = Path(__file__).parent / "new.docx"
        doc.save(output_path)

        self.popup("Document Generated", f"Document has been saved here: {output_path}")

    def on_ccode_select(self, spinner, text):
        
        self.root.ctitle_text = self.course_mapping.get(text, "")

    def on_tname_select(self, spinner, text):
        self.root.desg_text = self.teacher_mapping.get(text, "")

    def popup(self, title, message):
        popup = Popup(title=title, content=Label(text=message), size_hint=(None, None), size=(400, 200))
        popup.open()

if __name__ == "__main__":
    LayoutApp().run()
