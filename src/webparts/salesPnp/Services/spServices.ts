import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IDropdownOption } from "office-ui-fabric-react";
import { SPHttpClient, SPHttpClientResponse,IHttpClientOptions } from "@microsoft/sp-http";
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
    /**
     * Additems
     */
    public validateAndAdditems(context: WebPartContext, customerData: any, productData: any): Promise<string> {
        // Validation 
        let staus: string = "";
        let restApiUrl: string =
            context.pageContext.web.absoluteUrl +
            "/_api/web/lists/getByTitle('Orders')/items";
        const body: string = JSON.stringify({
            "Title" : "Order1",            
        });
        // "CustomerID": customerData,
        // "ProductID" : productData.ProductID,
        // "UnitsSold" : productData.NumberofUnits,
        // "UnitPrice" : productData.ProductUnitPrice,
        // "OrderStatus" : "Approved",
        const options: IHttpClientOptions = {
            body: body,
        };
        return new Promise<string>(async (resolve, reject) => {
            context.spHttpClient.post(restApiUrl, SPHttpClient.configurations.v1, options)
            .then((response: SPHttpClientResponse) => {
                console.log(response);
                if(response.ok){
                    response.json().then(
                        (result: any) => {
                            console.log(result);
                            resolve("Order with ID " + result.ID + " created Successfully!");
                        },
                        (error: any): void => {
                            reject("error occured while creating order!" + error);
                        }
                        );
                    }
                    else {
                        resolve("Order UnSuccessfully!");
                    }
                });
        });
    }

    /**
     * updateItem
     */
    public updateItem(context: WebPartContext, Orderid:number) {
        
    }
  
}