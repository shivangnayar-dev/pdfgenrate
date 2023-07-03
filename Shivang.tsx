import * as React from 'react';
import styles from './Shivang.module.scss';

import { escape } from '@microsoft/sp-lodash-subset';
import html2pdf from 'html2pdf.js';
import { IPdffProps } from './IPdffProps';

export default class Pdff extends React.Component<IPdffProps, {}> {
  generatePDF = () => {
    const element = document.querySelector('#pdf-content');
  
    if (element) {
      html2pdf().from(element).outputPdf('datauristring').then((pdfDataUri: string) => {
        const modalWidth = 800;
        const modalHeight = 600;
  
        // Create a new div element for the modal
        const modal = document.createElement('div');
        modal.style.width = `${modalWidth}px`;
        modal.style.height = `${modalHeight}px`;
        modal.style.position = 'fixed';
        modal.style.top = '50%';
        modal.style.left = '50%';
        modal.style.transform = 'translate(-50%, -50%)';
        modal.style.backgroundColor = '#ffffff';
        modal.style.zIndex = '9999';
  
        // Create an iframe element to display the PDF
        const iframe = document.createElement('iframe');
        iframe.style.width = '100%';
        iframe.style.height = '100%';
        iframe.src = pdfDataUri;
  
        // Create a close button
        const closeButton = document.createElement('button');
        closeButton.innerHTML = 'Close';
        closeButton.style.position = 'absolute';
        closeButton.style.top = '10px';
        closeButton.style.right = '10px';
        closeButton.style.padding = '5px';
  
        // Attach a click event listener to the close button
        closeButton.addEventListener('click', () => {
          // Remove the modal from the document
          document.body.removeChild(modal);
        });
  
        // Append the iframe and close button to the modal
        modal.appendChild(closeButton);
        modal.appendChild(iframe);
  
        // Append the modal to the document body
        document.body.appendChild(modal);
      });
    }
  };
   #thiswillgenratepdfinalternatewindow

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
      <section className={`${styles.shivang} ${hasTeamsContext ? styles.teams : ''}`}>
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
