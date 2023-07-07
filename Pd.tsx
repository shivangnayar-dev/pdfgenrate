import * as React from 'react';
import styles from './Pdff.module.scss';
import { uploadFileToLibrary} from './addattachement';

import { escape } from '@microsoft/sp-lodash-subset';
import html2pdf from 'html2pdf.js';
import { IPdffProps } from './IPdffProps';

export default class Pdff extends React.Component<IPdffProps, {}> {
  generatePDF = async () => {
    const element = document.querySelector('#pdf-content');
  
    if (element) {
      try {
        const pdfOptions = {
          output: 'save',
          filename: 'pdf-document.pdf' // Specify the desired file name
        };
  
        const pdfDataUri = await new Promise<string>((resolve, reject) => {
          html2pdf().set(pdfOptions).from(element).outputPdf('datauristring').then(resolve).catch(reject);
        });
  
        // Save the PDF file locally
        const downloadLink = document.createElement('a');
        downloadLink.href = pdfDataUri;
        downloadLink.download = 'pdf-document.pdf';
        downloadLink.click();
  
        // Convert the data URI to a Blob
        const blob = await (await fetch(pdfDataUri)).blob();
  
        // Create a File object from the Blob
        const pdfFile = new File([blob], 'pdf-document.pdf', { type: 'application/pdf' });
  
        // Upload the PDF file to the document library
        await uploadFileToLibrary(pdfFile.name, pdfFile);
  
        console.log('File saved locally and uploaded successfully.');
      } catch (error) {
        console.log(`Error generating PDF: ${error}`);
      }
    }
  };
  

   /*generatePDF = () => {
    const element = document.querySelector('#pdf-content');
  
    if (element) {
      const pdfOptions = {
        output: 'dataurlnewwindow'
      };
  
      html2pdf().from(element).set(pdfOptions).save();
    }
  };
*/
  
  
  
  public render(): React.ReactElement<IPdffProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <section className={`${styles.pdff} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.welcome}>
          <img alt="" src={isDarkTheme ? require('../assets/welcome-dark.png') : require('../assets/welcome-light.png')} className={styles.welcomeImage} />
          <h2>Well done, {escape(userDisplayName)}!</h2>
          <div>{environmentMessage}</div>
          <div>Web part property value: <strong>{escape(description)}</strong></div>
        </div>
        <div id="pdf-content">
          <h3>Welcome to SharePoint Framework!</h3>
          <p>
            The SharePoint Framework (SPFx) is an extensibility model for Microsoft Viva, Microsoft Teams, and SharePoint. It's the easiest way to extend Microsoft 365 with automatic Single Sign-On, automatic hosting, and industry-standard tooling.
          </p>
          <h4>Learn more about SPFx development:</h4>
          <ul className={styles.links}>
            <li><a href="https://aka.ms/spfx" target="_blank" rel="noreferrer">SharePoint Framework Overview</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-graph" target="_blank" rel="noreferrer">Use Microsoft Graph in your solution</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-teams" target="_blank" rel="noreferrer">Build for Microsoft Teams using SharePoint Framework</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-viva" target="_blank" rel="noreferrer">Build for Microsoft Viva Connections using SharePoint Framework</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-store" target="_blank" rel="noreferrer">Publish SharePoint Framework applications to the marketplace</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-api" target="_blank" rel="noreferrer">SharePoint Framework API reference</a></li>
            <li><a href="https://aka.ms/m365pnp" target="_blank" rel="noreferrer">Microsoft 365 Developer Community</a></li>
          </ul>
        </div>
        <button onClick={this.generatePDF}>Generate PDF</button>
      </section>
    );
  }
}