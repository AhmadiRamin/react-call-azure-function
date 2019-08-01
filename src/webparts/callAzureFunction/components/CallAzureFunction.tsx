import * as React from 'react';
import styles from './CallAzureFunction.module.scss';
import { ICallAzureFunctionProps } from './ICallAzureFunctionProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { AadHttpClient, HttpClientResponse, IHttpClientOptions } from '@microsoft/sp-http';

export default class CallAzureFunction extends React.Component<ICallAzureFunctionProps, {}> {
  public async componentDidMount(){
    const headers : Headers= new Headers();
    headers.append("Accept","application/json");
    const requestOptions : IHttpClientOptions={
      headers
    };

    const httpResponse:HttpClientResponse= await this.props.aadHttpClient.get(
      "https://spfx-test-app.azurewebsites.net/api/customers?code=d6fArLBGfTiin3iIvPmEzMM4BotFFzUNsTODCx8o0rnA4XdlskJa7Q==",
      AadHttpClient.configurations.v1,
      requestOptions      
      );
    const response = await httpResponse.json();
    console.log(response);
  }
  public render(): React.ReactElement<ICallAzureFunctionProps> {
    return (
      <div className={ styles.callAzureFunction }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <span className={ styles.title }>Welcome to SharePoint!</span>
              <p className={ styles.subTitle }>Customize SharePoint experiences using Web Parts.</p>
              <p className={ styles.description }>{escape(this.props.description)}</p>
              <a href="https://aka.ms/spfx" className={ styles.button }>
                <span className={ styles.label }>Learn more</span>
              </a>
            </div>
          </div>
        </div>
      </div>
    );
  }
}
