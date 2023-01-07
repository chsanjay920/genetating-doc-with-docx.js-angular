import { Component } from '@angular/core';
import * as docx from 'docx';
import { saveAs } from 'file-saver';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css'],
})
export class AppComponent {
  title = 'gdocx';
  status = 'not generated';
  img: any;

  public async createImageArrayBuffer() {
    return await fetch(
      '../assets/p.jpg'
    ).then((response) => {
      if (!response.ok) {
        throw new Error(`HTTP error, status = ${response.status}`);
      }
      return response.arrayBuffer();
    });
  }

  public generate(): void {
    this.img = this.createImageArrayBuffer().then((imahe) => {
      this.status = 'generating docx';
      const image = new docx.ImageRun({
        data: imahe,
        transformation: {
          width: 500,
          height: 500,
        },
      });
      
      const doc = new docx.Document({
        sections: [
          {
            properties: {},
            children: [
              new docx.Paragraph({
                children: [
                  image,
                  new docx.TextRun('Hello World'),
                  new docx.TextRun({
                    text: 'Foo Bar',
                    bold: true,
                  }),
                  new docx.TextRun({
                    text: '\tGithub is the best',
                    bold: true,
                  }),
                ],
              }),
            ],
          },
        ],
      });
      this.saveDocx(doc);
    });
  }

  public saveDocx(doc:any):void{
    docx.Packer.toBlob(doc).then((blob) => {
      saveAs(blob, 'example.docx');
      this.status= "docx generated sucessfully";
    });
  }
}
