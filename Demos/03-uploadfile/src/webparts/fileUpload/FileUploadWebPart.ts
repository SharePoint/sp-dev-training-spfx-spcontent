// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './FileUploadWebPart.module.scss';
import * as strings from 'FileUploadWebPartStrings';

import {
  ISPHttpClientOptions,
  SPHttpClient
} from '@microsoft/sp-http';

export interface IFileUploadWebPartProps {
  description: string;
}

export default class FileUploadWebPart extends BaseClientSideWebPart<IFileUploadWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${styles.fileUpload}">
        <div class="${styles.container}">
          <div class="${styles.row}">
            <div class="${styles.column}">
              <input class="${styles.fileUpload}-fileUpload" type="file" /><br />
              <input class="${styles.fileUpload}-uploadButton" type="button" value="Upload" />
            </div>
          </div>
        </div>
      </div>`;

    // get reference to file control
    const inputFileElement = document.getElementsByClassName(`${styles.fileUpload}-fileUpload`)[0] as HTMLInputElement;

    // wire up button control
    const uploadButton = document.getElementsByClassName(`${styles.fileUpload}-uploadButton`)[0] as HTMLButtonElement;

    uploadButton.addEventListener('click', async () => {
      // get filename
      const filePathParts = inputFileElement.value.split('\\');
      const fileName = filePathParts[filePathParts.length - 1];

      // get file data
      const fileData = await this._getFileBuffer(inputFileElement.files[0]);

      // upload file
      await this._uploadFile(fileData, fileName);
    });
  }

  private _getFileBuffer(file: File): Promise<ArrayBuffer> {
    return new Promise((resolve, reject) => {
      let fileReader = new FileReader();

      // write up error handler
      fileReader.onerror = (event: ProgressEvent<FileReader>) => {
        reject(event.target.error);
      };

      // wire up when finished reading file
      fileReader.onloadend = (event: ProgressEvent<FileReader>) => {
        resolve(event.target.result as ArrayBuffer);
      };

      // read file
      fileReader.readAsArrayBuffer(file);

    });
  }

  private async _uploadFile(fileData: ArrayBuffer, fileName: string): Promise<void> {

    // create target endpoint for REST API HTTP POST
    const endpoint = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('Documents')/RootFolder/Files/add(overwrite=true,url='${fileName}')`;

    const options: ISPHttpClientOptions = {
      headers: { 'CONTENT-LENGTH': fileData.byteLength.toString() },
      body: fileData
    };

    // upload file
    const response = await this.context.spHttpClient.post(endpoint, SPHttpClient.configurations.v1, options);

    if (response.status === 200) {
      alert('File uploaded successfully');
    } else {
      throw new Error(`Error uploading file: ${response.statusText}`);
    }
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
