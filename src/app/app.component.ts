import { ChangeDetectorRef, Component, type AfterViewInit } from '@angular/core';
import { CKEditorModule } from '@ckeditor/ckeditor5-angular';
import { type EditorConfig, InlineEditor, AutoLink, Autosave, Bold, Essentials, Italic, Link, Paragraph } from 'ckeditor5';
import { saveAs } from 'file-saver';
import * as docx from 'docx';
import jsPDF from 'jspdf';
import html2canvas from 'html2canvas';
import { CommonModule } from '@angular/common';
const LICENSE_KEY = 'GPL'; // or <YOUR_LICENSE_KEY>.
@Component({
  selector: 'app-root',
  standalone: true,
  imports: [CommonModule, CKEditorModule],
  templateUrl: './app.component.html',
  styleUrl: './app.component.css',
  template: `
    <ckeditor [editor]="Editor" [(ngModel)]="editorData"></ckeditor>
  `,
})
export class AppComponent implements AfterViewInit {
  title = 'test-pcc';
  constructor(private changeDetector: ChangeDetectorRef) { }

  public isLayoutReady = false;
  public Editor = InlineEditor;
  public config: EditorConfig = {}; // CKEditor needs the DOM tree before calculating the configuration.
  public ngAfterViewInit(): void {
    this.config = {
      plugins: [AutoLink, Autosave, Bold, Essentials, Italic, Link, Paragraph],
      initialData:
        `
    <div style="font-family:'TH Sarabun New', sans-serif; font-size:18pt; line-height:1.8;">
      <p style="text-align:right;">ที่ ศธ ........../..........</p>
      <p style="text-align:right;">กรมสรรพสามิต</p>
      <p style="text-align:right;">ถนนนครไชยศรี กรุงเทพฯ 10300</p>
      <p style="text-align:right;">๒๐ กรกฎาคม ๒๕๖๔</p>

      <p><strong>เรื่อง</strong> แจ้งให้ไปทำสัญญาจ้าง............................................................</p>
      <p><strong>เรียน</strong> กรรมการผู้จัดการ............................................................</p>
      <p>อ้างถึง ใบเสนอราคาด้วยวิธี.................................. เลขที่................... ลงวันที่....................</p>
      <p>ตามที่อ้างถึง ..................................................... ได้เสนอราคางานโครงการ................................................</p>

      <p>เป็นเงินทั้งสิ้น ........................................ บาท</p>

      <p>จึงแจ้งให้ไปทำสัญญาภายใน ....... วันทำการ นับถัดจากวันที่ได้รับหนังสือนี้</p>
      <p>โดยนำหลักประกันสัญญาจำนวน ร้อยละ ........... เป็นเงิน ............. บาท (.............................) ไปวางเป็นหลักประกัน</p>

      <p>หากไม่ทำสัญญาตามเวลาที่กำหนด กรมฯ ขอสงวนสิทธิ์เรียกร้องค่าเสียหายอื่น (ถ้ามี)</p>

      <p style="text-align:center;">ขอแสดงความนับถือ</p>
      <br><br>
      <p style="text-align:center;">(................................................)</p>
      <p style="text-align:center;">ผู้อำนวยการสำนักงานบริหารการคลังและรายได้</p>
      <p style="text-align:center;">ปฏิบัติราชการแทน</p>
      <p style="text-align:center;">อธิบดีกรมสรรพสามิต</p>
    </div>
  `,
      licenseKey: LICENSE_KEY,
      link: {
        addTargetToExternalLinks: true,
        defaultProtocol: 'https://',
        decorators: {
          toggleDownloadable: {
            mode: 'manual',
            label: 'Downloadable',
            attributes: {
              download: 'file'
            }
          }
        }
      },
      placeholder: 'Type or paste your content here!'
    };

    this.isLayoutReady = true;
    this.changeDetector.detectChanges();
  }


  downloadFromConfig(type: 'pdf' | 'docx') {
    const html = this.config.initialData;

    if (type === 'docx') {
      this.exportToWord(html);
    } else if (type === 'pdf') {
      this.exportToPDF(html);
    }
  }

  exportToWord(html: any) {
    const { Document, Packer, Paragraph } = docx;
    const textOnly = this.stripHtml(html);

    const doc = new Document({
      sections: [{
        children: [
          new Paragraph({
            children: [
              new docx.TextRun({
                text: textOnly,
                font: 'TH Sarabun New',
                size: 32
              })
            ]
          })
        ]
      }]
    });

    Packer.toBlob(doc).then(blob => {
      const filename = this.getThaiDateFilename('หนังสือราชการ', 'docx');
      saveAs(blob, filename);
    });
  }

  exportToPDF(html: any) {
    const container = document.createElement('div');
    container.innerHTML = html;
    container.style.width = '794px'; // A4 width
    container.style.padding = '20px';
    container.style.fontFamily = "'TH Sarabun New', sans-serif";
    document.body.appendChild(container);

    html2canvas(container).then(canvas => {
      const imgData = canvas.toDataURL('image/png');
      const pdf = new jsPDF('p', 'mm', 'a4');
      const pdfWidth = 210;
      const pdfHeight = (canvas.height * pdfWidth) / canvas.width;
      pdf.addImage(imgData, 'PNG', 0, 0, pdfWidth, pdfHeight);
      const filename = this.getThaiDateFilename('หนังสือราชการ', 'docx');
      pdf.save(filename);
      document.body.removeChild(container); // cleanup
    });
  }

  stripHtml(html: string): string {
    const div = document.createElement('div');
    div.innerHTML = html;
    return div.textContent || div.innerText || '';
  }

  getThaiDateFilename(prefix: string, ext: string): string {
    const now = new Date();
    const day = String(now.getDate()).padStart(2, '0');
    const month = String(now.getMonth() + 1).padStart(2, '0'); // 0-index
    const year = now.getFullYear();
    return `${prefix}_${day}-${month}-${year}.${ext}`;
  }
}
