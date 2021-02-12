import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IDropdownOption } from "office-ui-fabric-react";
import { SPHttpClient, SPHttpClientResponse, IHttpClientOptions } from "@microsoft/sp-http";
import { sp } from "@pnp/sp/presets/all";
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
     * getOrderList
     */
    public getOrderList(context: WebPartContext): Promise<IDropdownOption[]> {
        let orderIdList: IDropdownOption[] = [];
        let restApiurl: string = context.pageContext.web.absoluteUrl + "/_api/web/lists/getbytitle('Orders')/items?select=Id";
        return new Promise<IDropdownOption[]>(async (resolve, reject) => {
            context.spHttpClient
                .get(restApiurl, SPHttpClient.configurations.v1)
                .then((response: SPHttpClientResponse) => {
                    response.json().then((results: any) => {
                        // console.log(results);
                        results.value.map((result: any) => {
                            // console.log(result.CustomerName);
                            orderIdList.push({
                                key: result.ID,
                                text: result.Id
                            });
                        });
                    });
                    resolve(orderIdList);
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
                            // console.log(result.ProductName);
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
    public createItems(context: WebPartContext, state: any): Promise<string> {
        // Validation 
        let staus: string = "";
        let restApiUrl: string =
            context.pageContext.web.absoluteUrl +
            "/_api/web/lists/getByTitle('Orders')/items";
        // console.log(customerData.CustomerId,productData.ProductId);
        const body: string = JSON.stringify({
            Customer_x0020_IDId: state.CustomerId,
            Product_x0020_IDId: state.ProductId,
            UnitsSold: state.NumberofUnits,
            SaleValue: state.TotalValue,
            OrderStatus: "Approved",
        });
        // console.log(body);
        const options: IHttpClientOptions = {
            headers: {
                Accept: "application/json;odata=nometadata",
                "content-type": "application/json;odata=nometadata",
                "odata-version": "",
            },
            body: body,
        };
        return new Promise<string>(async (resolve, reject) => {
            context.spHttpClient.post(restApiUrl, SPHttpClient.configurations.v1, options)
                .then((response: SPHttpClientResponse) => {
                    // console.log(response);
                    if (response.ok) {
                        response.json().then(
                            (result: any) => {
                                // console.log(result);
                                resolve("Order with ID: " + result.Id + " created Successfully!");
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
     * getUpdateitem
     */
    public getUpdateitem(context: WebPartContext, data: any): Promise<any> {
        console.log("getUpdateitem Called!");
        let restApiUrl: string =
            context.pageContext.web.absoluteUrl +
            "/_api/web/lists/getbytitle('Orders')/items?$filter=(ID eq " + data.text + ")";

        return new Promise<any>(async (resolve, reject) => {
            context.spHttpClient
                .get(restApiUrl, SPHttpClient.configurations.v1)
                .then((response: SPHttpClientResponse) => {
                    console.log(response);
                    if (response.ok) {
                        response.json().then((results) => {
                            console.log(results);
                            resolve(results);
                        });
                    }
                }, (error: any) => {
                    reject("getUpdateitem failed");
                });
        });

    }


    /**
     * getProductDetails
     */
    public getProductDetails(context: WebPartContext, data: any) {
        let restApiUrl: string =
            context.pageContext.web.absoluteUrl +
            "/_api/web/lists/getbytitle('Products')/items?$filter=(ID eq " +
            data.key +
            ") and (ProductName eq '" +
            data.text +
            "')";

        return new Promise<any>(async (resolve, reject) => {
            context.spHttpClient.get(restApiUrl, SPHttpClient.configurations.v1)
                .then((response: SPHttpClientResponse) => {
                    if (response.ok) {
                        response.json().then((results: any) => {
                            resolve(results.value[0]);
                        });
                    }
                }, (error: any) => {
                    reject("getUpdateitem failed");
                });
        });
    }

    /**
     * updateItem
     */
    public updateItem(context: WebPartContext, state: any) {
        console.log("updateItem Called!");
        // Upload modified data to Order list
        let restApiUrl: string =
            context.pageContext.web.absoluteUrl +
            "/_api/web/lists/getByTitle('Orders')/items('" + state.orderId + "')";
        // console.log(customerData.CustomerId,productData.ProductId);
        const body: string = JSON.stringify({
            Customer_x0020_IDId: state.CustomerId,
            Product_x0020_IDId: state.ProductId,
            UnitsSold: state.NumberofUnits,
            SaleValue: state.TotalValue,
            OrderStatus: "Approved",
        });
        // console.log(body);
        const options: IHttpClientOptions = {
            headers: {
                Accept: "application/json;odata=nometadata",
                "content-type": "application/json;odata=nometadata",
                "odata-version": "",
                "IF-MATCH": "*",
                "x-HTTP-METHOD": "MERGE",
            },
            body: body,
        };
        return new Promise<string>(async (resolve, reject) => {
            context.spHttpClient.post(restApiUrl, SPHttpClient.configurations.v1, options)
                .then((response: SPHttpClientResponse) => {
                    // console.log(response);
                    if (response.ok) {
                        resolve("Order Updated Successfully!");
                    }
                    else {
                        resolve("Order Update Unsuccesful!");
                    }
                },
                    (error: any): void => {
                        reject("error occured while creating order!" + error);
                    });
        });
    }
    /**
     * deleteItem(this.props.context, this.state.orderId)
     */
    public deleteItem(context: WebPartContext, orderId: any) {
        console.log("deleteItem Called!");
        let restApiUrl: string =
            context.pageContext.web.absoluteUrl +
            "/_api/web/lists/getByTitle('Orders')/items('" + orderId + "')";
            //?$filter=(ID eq "+ parseInt(orderId) + ")";
            
        const options: IHttpClientOptions = {
            headers: {
                Accept: "application/json;odata=nometadata",
                "content-type": "application/json;odata=nometadata",
                "odata-version": "",
                "IF-MATCH": "*",
                "X-HTTP-METHOD": "DELETE",
            }
        };
        
        return new Promise<string>(async(resolve,reject) => {
            context.httpClient.post(restApiUrl, SPHttpClient.configurations.v1,options)
            .then((response : SPHttpClientResponse) => {
                resolve("Order Id: " + orderId + "Delete!");
            },(error:any) =>{
                reject("Delete not successful!");
            });
        });
    }

    /**
   * deleteItem2
   */
  public deleteItem2 = async () =>{
    console.log("deleteItem2 Called!");
    // let list = await sp.web.getList("/sites/Jaguar/lists/Orders").items.getById(13).recycle()
    let list = await sp.web.getList("/sites/Jaguar/lists/Orders");
    console.log(list);
  }

}