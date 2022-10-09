import { Component, OnInit } from '@angular/core';
import { MsalService } from '@azure/msal-angular';

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

  constructor(private msalService: MsalService) {}

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

  dialog;
  openLoginPopup(): void {
    //Office.context.ui.displayDialogAsync('https://login.microsoftonline.com/0c0cc3c4-4b87-4a2a-8003-1aa4656d1f0a/oauth2/v2.0/authorize?client_id=bccf936a-d9b3-4ced-91f1-871ffbedb83a&scope=user.read',
    //Office.context.ui.openBrowserWindow('https://localhost:4200#/login');
    //Office.context.ui.displayDialogAsync('https://login.microsoftonline.com/0c0cc3c4-4b87-4a2a-8003-1aa4656d1f0a/oauth2/v2.0/authorize?client_id=bccf936a-d9b3-4ced-91f1-871ffbedb83a&scope=user.read%20openid%20profile%20offline_access&redirect_uri=https%3A%2F%2Flocalhost%3A4200&client-request-id=d7b0d807-d940-4830-b8fd-bf6242c851e3&response_mode=fragment&response_type=code&x-client-SKU=msal.js.browser&x-client-VER=2.28.3&client_info=1&code_challenge=4ODX0JdXsC2oRwRDx9i0zWJGuGSVnUWbzX0otLjGLu0&code_challenge_method=S256&nonce=8dda236c-b824-40af-890d-10f01abd207b&state=eyJpZCI6IjJjOWIzZDY3LWJlMWYtNDNiMy1iMTRiLWU5ZWVjOTc4ODAwYSIsIm1ldGEiOnsiaW50ZXJhY3Rpb25UeXBlIjoicG9wdXAifX0%3D',
    Office.context.ui.displayDialogAsync('https://localhost:4200#/login',
    { height: 40, width: 30, displayInIframe: true },
      (result) => {
        this.dialog = result.value;
        this.dialog.addEventHandler(
          Office.EventType.DialogMessageReceived,
          this.processMessage
        );
      }
    );
  }

  processMessage(arg) {
    var messageFromDialog = JSON.parse(arg.message);
    console.log(messageFromDialog.name);
    this.dialog.close();
  }
  isUserLoggedIn(): boolean {
    return this.msalService.instance.getActiveAccount() !== null;
  }
}
