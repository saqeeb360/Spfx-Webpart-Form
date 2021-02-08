import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IDropdownOption } from "office-ui-fabric-react";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
export class spOperation {
    /**
     * getCustomerNameList
     * Using rest calls
     */
    public getCustomerNameList(context: WebPartContext): Promise<IDropdownOption[]> {
        let customerNameList: IDropdownOption[] = [];
        let restApiurl: string = context.pageContext.web.absoluteUrl + "/_api/web/lists/getbytitle('Customers')/items?select=CustomerName";
        return new Promise<IDropdownOption[]>(async (resolve, reject) => {
            context.spHttpClient
                .get(restApiurl, SPHttpClient.configurations.v1)
                .then((response: SPHttpClientResponse) => {
                    response.json().then((results: any) => {
                        // console.log(results);
                        results.value.map((result: any) => {
                            // console.log(result.CustomerName);
                            customerNameList.push({
                                key: result.ID,
                                text: result.CustomerName
                            });
                        });
                    });
                    resolve(customerNameList);
                }, (error: any) => {
                    reject("error occured in getListTitle() ");
                });
        });
    }

    /**
     * getProductNameList
     */
    public getProductNameList(context: WebPartContext): Promise<IDropdownOption[]> {
        let productNameList: IDropdownOption[] = [];
        let restApiurl: string = context.pageContext.web.absoluteUrl + "/_api/web/lists/getbytitle('Products')/items";
        return new Promise<IDropdownOption[]>(async (resolve, reject) => {
            context.spHttpClient
                .get(restApiurl, SPHttpClient.configurations.v1)
                .then((response: SPHttpClientResponse) => {
                    response.json().then((results: any) => {
                        console.log(results);
                        results.value.map((result: any) => {
                            console.log(result.ProductName);
                            productNameList.push({
                                key: result.ID,
                                text: result.ProductName,
                                // data: {
                                //     ProductType: result.ProductType,
                                //     Product_x0020_Unit_x0020_Price:
                                //         result.Product_x0020_Unit_x0020_Price,
                                //     ProductExpiryDate: result.ProductExpiryDate
                                // }
                            });
                        });
                    });
                    resolve(productNameList);
                }, (error: any) => {
                    reject("error occured in getListTitle() ");
                });
        });
    }
}