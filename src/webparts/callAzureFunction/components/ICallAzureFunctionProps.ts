import { AadHttpClient } from "@microsoft/sp-http";

export interface ICallAzureFunctionProps {
  description: string;
  aadHttpClient: AadHttpClient;
}
