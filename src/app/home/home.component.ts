import { Component, OnInit } from '@angular/core';

declare const Office: any;
declare const OfficeExtension: any;
declare const Word: any;
declare const $: any;

@Component({
  selector: 'app-home',
  templateUrl: './home.component.html',
  styleUrls: ['./home.component.scss'],
})
export class HomeComponent implements OnInit {
  public selectedText: string = '';
  public replaceText: string = '';

  constructor() {}

  ngOnInit(): void {}

  showSelectedText() {
    this.getSelectedText((selectedText: string) => {
      this.selectedText = selectedText;
    });
  }

  // Mark section with content control
  getSelectedText(callback: any): any {
    
    // Run a batch operation against the Word object model.
    Word.run((context: any) => {
      // Create a proxy range object for the current selection.
      var range = context.document.getSelection();
      range.load();
      return context.sync().then(async () => {
        callback(range.text);
      });
    }).catch((error) => {
      console.log('Error: ' + JSON.stringify(error));
      if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
        callback('');
      }
    });
  }

  replaceDocumentText() {
    this.lockUnlockParentControl(false, (result: boolean) => {
      if (result) {
        this.updateAuditTable((rowInserted: boolean) => {
          if (rowInserted) {
            this.replaceTextInDocument(
              this.replaceText,
              (textInserted: boolean) => {
                if (textInserted) {
                  this.lockUnlockParentControl(true, (result: boolean) => {});
                }
              }
            );
          }
        });
      }
    });
  }
  replaceTextInDocument(text: string, callback: any): any {
    let instance = this;
    // Run a batch operation against the Word object model.
    Word.run((context: any) => {
      // Create a proxy range object for the current selection.
      var range = context.document.getSelection();
      range.insertText(text, Word.InsertLocation.replace);
      return context.sync().then(async () => {
        callback(true);
      });
    }).catch((error) => {
      console.log('Error: ' + JSON.stringify(error));
      if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
        callback(false);
      }
    });
  }

  textChanged(event) {
    this.replaceText = event.target.value;
  }
  lockUnlockParentControl(lock: boolean, callback: any): any {
    let instance = this;
    // Run a batch operation against the Word object model.
    Word.run((context: any) => {
      // Create a proxy range object for the current selection.
      var ctrl = context.document.getSelection().parentContentControl;
      ctrl.load('tag, cannotDelete, cannotEdit');
      return context.sync().then(async () => {
        let parentMarker: string = ctrl.tag;
        if (parentMarker.startsWith('parent')) {
          console.log(
            'Parent control found [' +
              ctrl.tag +
              '] with cannotEdit: ' +
              ctrl.cannotEdit +
              ' cannotEdit: ' +
              ctrl.cannotDelete
          );
          ctrl.cannotEdit = lock;
          ctrl.cannotDelete = lock;
          return context.sync().then(async () => {
            console.log('Parent control is locked now [' + lock + ']');
            callback(true);
          });
        } else {
          console.log('Parent control not found');
          callback(false);
        }
        callback(true);
      });
    }).catch((error) => {
      console.log('Error: ' + JSON.stringify(error));
      if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
        callback(false);
      }
    });
  }

  updateAuditTable(callback: any) {
    Word.run(async (context: any) => {
      var contentControl = context.document.contentControls
        .getByTag('changeLogTable')
        .getFirst();
      var table = contentControl.tables.getFirst();

      context.load(table);
      return context.sync().then(async () => {
        var rowCount: number = table.rowCount;
        var newRowIndex: number = rowCount;
        table.addRows('End', 1);
        table.getCell(newRowIndex, 0).value = 'User ' + newRowIndex;
        table.getCell(newRowIndex, 1).value =
          'Document updated ' + newRowIndex + '';
        table.getCell(newRowIndex, 2).value = new Date().toISOString();
        return context.sync().then(async () => {
          console.log('inserted values.');
          callback(true);
        });
      });
    }).catch(async (error: any) => {
      console.log('Error: ' + JSON.stringify(error));
      if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
      }
      callback(false);
    });
  }

   
  openLoginPopup(): void {    
    Office.context.ui.displayDialogAsync('https://localhost:4200/login',
      { height: 40, width: 30,displayInIframe: false});
  }
 
}
