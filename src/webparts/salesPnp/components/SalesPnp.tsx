import * as React from "react";
import styles from "./SalesPnp.module.scss";
import { ISalesPnpProps } from "./ISalesPnpProps";
import { ISalesPnpState } from "./ISalesPnpState";
import { escape } from "@microsoft/sp-lodash-subset";
import { Dropdown, TextField } from "office-ui-fabric-react";
import { spOperation } from "../Services/spServices";
import {SPHttpClient , SPHttpClientResponse} from "@microsoft/sp-http";

export default class SalesPnp extends React.Component<
  ISalesPnpProps,
  ISalesPnpState,
  {}
> {
  public _spOps: spOperation;
  constructor(props: ISalesPnpProps) {
    super(props);
    this._spOps = new spOperation();
    this.state = {
      customerNameList: [],
      productNameList : [],
      customerData: {
        CustomerName: "",
        CustomerId: "",
      },
      productData: {
        ProductId: "",
        ProductName: "",
        ProductUnitPrice: "",
        ProductExpiryDate: "",
        ProductType: "",
        NumberofUnits : "",
        TotalValue : ""
      },
    };
  }
  public componentDidMount() {
    // getCustomerNameList
    this._spOps.getCustomerNameList(this.props.context).then((result: any) => {
      this.setState({ customerNameList: result });
      // console.log(this.state.customerNameList);
    });

    // getProductNameList
    this._spOps.getProductNameList(this.props.context).then((result: any) => {
      this.setState({ productNameList: result });
      // console.log(this.state.productNameList);
    });
    }

  /**
   * getCustomerName
   * this function is called when a dropdown item is changed
   * To save the CustomerName and id to this.state
   */
  public getCustomerName = (event: any, data: any) => {
    // console.log(data);
    this.setState({customerData : {CustomerName : data.text,CustomerId : data.key}});
  }

  /**
   * getProductName
   * this function is called when a dropdown item is changed
   * To save the productName,id,type,date and unit value to this.state
   * and show it in the form automatically
   */
  public getProductName = (event: any, data: any) => {
    // console.log(data);
    let restApiUrl:string = this.props.context.pageContext.web.absoluteUrl + "/_api/web/lists/getbytitle('Products')/items?$filter=(ID eq "+data.key+") and (ProductName eq '" + data.text + "')";

    this.props.context.spHttpClient.get(restApiUrl,SPHttpClient.configurations.v1)
    .then((response:SPHttpClientResponse) =>{
      response.json().then((results:any) => {
        console.log(results);
        let result:any = results.value[0];
        let date = new Date(result.ProductExpiryDate);
        
        console.log(result.ProductExpiryDate);
        console.log(date);

        this.setState({
          productData : {
            ProductName : data.text,
            ProductId : data.key,
            ProductExpiryDate : date,
            ProductType : result.ProductType,
            ProductUnitPrice : result.Product_x0020_Unit_x0020_Price,
            NumberofUnits : this.state.productData.NumberofUnits,
            TotalValue : this.state.productData.TotalValue,
          }
        });
      });
    });
  }

  /**
   * setNumberofUnits
   */
  public setNumberofUnits = (event:any, data:any) =>{
    if(this.state.productData.ProductUnitPrice !== ""){
      let numberofUnits : number = parseInt(data);
      let priceofunit :number = parseInt(this.state.productData.ProductUnitPrice); 
      let totalValue : number = numberofUnits * priceofunit ;
      // console.log(numberofUnits);
      // console.log(priceofunit);
      // console.log(totalValue);
      this.setState({
        productData : {
          NumberofUnits : numberofUnits,
          TotalValue : totalValue,
          ProductId : this.state.productData.ProductId,
          ProductName : this.state.productData.ProductName,
          ProductExpiryDate :this.state.productData.ProductExpiryDate ,
          ProductType : this.state.productData.ProductType,
          ProductUnitPrice : this.state.productData.ProductUnitPrice,
        }
      });
    }
    return ;
  }

  public render(): React.ReactElement<ISalesPnpProps> {
    return (
      <div className={styles.salesPnp}>
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.column}>
              <span className={styles.title}>Welcome to SharePoint!</span>
              <p className={styles.subTitle}>
                Customize SharePoint experiences using Web Parts.
              </p>
              <p className={styles.description}>
                {escape(this.props.description)}
              </p>
            </div>
            <div className={styles.column}>
              Enter Customer Name
              <Dropdown
                required
                options={this.state.customerNameList}
                onChange={this.getCustomerName}
              ></Dropdown>
            </div>
            
            <div className={styles.column}>
              Enter Product Name
              <Dropdown
                required
                options={this.state.productNameList}
                onChange={this.getProductName}
              ></Dropdown>
            </div>
            
            <div className={styles.column}>
              <div>
                Product Type
                <TextField ariaLabel="disabled Product type" disabled placeholder={this.state.productData.ProductType} />
              </div>

              <div>
                Product Expiry Date
                <TextField ariaLabel="disabled Product type" disabled placeholder={this.state.productData.ProductExpiryDate === "" ? "" : new Date(this.state.productData.ProductExpiryDate).toDateString()} />
              </div>              

              <div>
                Product Unit Price
                <TextField ariaLabel="disabled Product type" disabled placeholder={this.state.productData.ProductUnitPrice} />
              </div>

              <div>
                Number of Units
                <TextField ariaLabel="required number" type="number" required
                onChange={this.setNumberofUnits} />
              </div>

              <div>
                Total Sales Price
                <TextField ariaLabel="disabled Product Sales Price" disabled placeholder={this.state.productData.TotalValue} />
              </div>

            </div>
          </div>
        </div>
      </div>
    );
  }
}
