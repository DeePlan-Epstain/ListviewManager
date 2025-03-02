import { ListViewCommandSetContext } from "@microsoft/sp-listview-extensibility";
import { SPHttpClientResponse, ISPHttpClientOptions, SPHttpClient } from "@microsoft/sp-http";
import axios from "axios";

// export enum PaUrls {
//     CONVERT_TO_PDF = 'https://prod-48.westeurope.logic.azure.com:443/workflows/60282bd80e29428c9094b301317c665c/triggers/manual/paths/invoke?api-version=2016-06-01'
// }

export default class PAService {
    private _context: ListViewCommandSetContext;
    public CONVERT_TO_PDF: string;

    constructor(private context: ListViewCommandSetContext, url: string) {
        this._context = this.context;
        this.CONVERT_TO_PDF = url
    };

    private async getAccessToken(): Promise<string> {
        const body: ISPHttpClientOptions = {
            body: JSON.stringify({
                resource: "https://service.flow.microsoft.com/"
            })
        };

        const token: any = await this._context.spHttpClient.post(
            `${this._context.pageContext.web.absoluteUrl}/_api/SP.OAuth.Token/Acquire`,
            SPHttpClient.configurations.v1 as any,
            body
        );

        const tokenJson = await token.json();

        return tokenJson.access_token;
    };

    public async get(url: string) {
        const token = await this.getAccessToken();

        try {
            const { data } = await axios.get(url,
                {
                    headers: {
                        'Content-Type': 'application/json',
                        'Authorization': `Bearer ${token}`
                    }
                }
            );

            return Promise.resolve(data);
        } catch (err) {
            return Promise.reject(err);
        }
    }

    public async post(url: string, body: any) {
        const token = await this.getAccessToken();

        try {
            const res = await axios.post(url, body, {
                headers: {
                    'Content-Type': 'application/json',
                    'Authorization': `Bearer ${token}`
                }
            });

            return Promise.resolve(res);
        } catch (err) {
            return Promise.reject(err);
        }
    }
}