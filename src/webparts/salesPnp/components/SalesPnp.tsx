import * as React from "react";
import styles from "./SalesPnp.module.scss";
import { ISalesPnpProps } from "./ISalesPnpProps";
import { ISalesPnpState } from "./ISalesPnpState";
import { escape, times } from "@microsoft/sp-lodash-subset";

import { spOperation } from "../Services/spServices";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";

import {
  DefaultButton,
  Dropdown,
  TextField,
  Stack,
  IStackTokens,
  IDropdownStyles,
} from "office-ui-fabric-react";
import { Label } from "office-ui-fabric-react/lib/Label";
import {
  Pivot,
  PivotItem,
  PivotLinkSize,
} from "office-ui-fabric-react/lib/Pivot";

const stackTokens: IStackTokens = { childrenGap: 100 };
const bigVertStack: IStackTokens = { childrenGap: 20 };
const SmallVertStack: IStackTokens = { childrenGap: 20 };
const dropdownStyles: Partial<IDropdownStyles> = { dropdown: { width: 250 } };

export default class SalesPnp extends React.Component<
  ISalesPnpProps,
  ISalesPnpState,
  {}
> {
  public _spOps: spOperation;
  constructor(props: ISalesPnpProps) {
    super(props);
    console.log("Constructor called!!");
    this._spOps = new spOperation();
    this.state = {
      customerNameList: [],
      productNameList: [],
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
        NumberofUnits: "",
        TotalValue: "",
      },
      orderIdList : [],
      orderId : "",
      status: "This is working",
    };
  }
  public componentDidMount() {
    console.log("Component Did Mount called!!");
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

    // getOrderList
    this._spOps.getOrderList(this.props.context).then((result:any) => {
      this.setState({orderIdList : result});
    })
  }

  /**
   * getCustomerName
   * this function is called when a dropdown item is changed
   * To save the CustomerName and id to this.state
   */
  public getCustomerName = (event: any, data: any) => {
    console.log("getCustomerName called!!");
    // console.log(data);
    this.setState({
      customerData: { CustomerName: data.text, CustomerId: data.key },
    });
  };

  /**
   * getProductName
   * this function is called when a dropdown item is changed
   * To save the productName,id,type,date and unit value to this.state
   * and show it in the form automatically
   */
  public getProductName = (event: any, data: any) => {
    console.log("getProductName called!!");
    // console.log(data);
    let restApiUrl: string =
      this.props.context.pageContext.web.absoluteUrl +
      "/_api/web/lists/getbytitle('Products')/items?$filter=(ID eq " +
      data.key +
      ") and (ProductName eq '" +
      data.text +
      "')";

    this.props.context.spHttpClient
      .get(restApiUrl, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        response.json().then((results: any) => {
          let result: any = results.value[0];
          let date = new Date(result.ProductExpiryDate);
          var totalValue:any;
          if(this.state.productData.NumberofUnits === ""){
            // Update Total Value when Number of Unit is not zero! otherwise don't update!!
            totalValue = this.state.productData.TotalValue;
          }
          else if(this.state.productData.NumberofUnits === "0"){
            totalValue = this.state.productData.TotalValue;
          }
          else{
            totalValue = result.Product_x0020_Unit_x0020_Price * this.state.productData.NumberofUnits;
          }
          // console.log(results);
          // console.log(result.ProductExpiryDate);
          // console.log(date);
          this.setState({
            productData: {
              ProductName: data.text,
              ProductId: data.key,
              ProductExpiryDate: date,
              ProductType: result.ProductType,
              ProductUnitPrice: result.Product_x0020_Unit_x0020_Price,
              NumberofUnits: this.state.productData.NumberofUnits,
              TotalValue: totalValue,
            },
          });
        });
      });
  };

  /**
   * setNumberofUnits
   */
  public setNumberofUnits = (event: any, data: any) => {
    console.log("setNumberofUnits called!!");
    console.log(this.state.productData.ProductUnitPrice, data);
    var numberofUnits:any;
    var totalValue:any;
    if(data === "0") {
      numberofUnits = data;
      totalValue = "";
    }
    else if(data===""){
      console.log("setNumberofUnits called -> In ifelse -> data = '' statement!!");
      numberofUnits = data;
      totalValue = "";
    }
    else if (this.state.productData.ProductUnitPrice === "") {
      console.log("setNumberofUnits called -> In ifelse -> UnitPrice = '' statement!!");
      numberofUnits = parseInt(data);
      totalValue = "";
    }
    else{
      console.log("setNumberofUnits called -> In else statement!!");
      numberofUnits = parseInt(data);
      var priceofunit: number = parseInt(
        this.state.productData.ProductUnitPrice
      );
      totalValue = numberofUnits * priceofunit;
    }
    
      // console.log(numberofUnits);
      // console.log(priceofunit);
      // console.log(totalValue);
      this.setState({
        productData: {
          NumberofUnits: numberofUnits,
          TotalValue: totalValue,
          ProductId: this.state.productData.ProductId,
          ProductName: this.state.productData.ProductName,
          ProductExpiryDate: this.state.productData.ProductExpiryDate,
          ProductType: this.state.productData.ProductType,
          ProductUnitPrice: this.state.productData.ProductUnitPrice,
        },
      });
    return;
  };
  /**
   * validateItem
   */
  public validateItem = () =>{
    console.log("ValidateItem called!!");
    let s1 = this.state.customerData;
    var status : any;
    for (let key in s1) {
      if(s1[key] === ""){
        console.log(key, s1[key],typeof s1[key]);
        status = "Wrong " + key +"!";
        this.setState({status : status});
        return;
      }
    }
    let s2 = this.state.productData;
    for (let key in s2) {
      if(s2[key] === ""){
        console.log(key, s2[key],typeof s2[key]);
        status = "Wrong " + key +"!";
        this.setState({status : status});
        return;
      }
    }
    console.log("Validate Complete and Uploading Order Details");
    this._spOps.createItems(this.props.context,this.state.customerData,this.state.productData)
    .then((result:string) =>{
      this.setState({status:result});
    });
  }
  /**
   * valUpdateitem
   */
  public getOrderDetailsToUpdate = (event:any,data:any) => {
    // Valid Order Id -> Not empty -> not zero
    console.log("getOrderDetailsToUpdate called!");
    if(data === ""){
      console.log("Empty data");
      return;
    }
    // Now get the data from rest and call setstate to change the state
    console.log(data);
    console.log(this.state.orderIdList);
    this._spOps.getUpdateitem(this.props.context,data)
    .then((result) => {
      console.log(result);
    });
  }

  /**
     * resetForm
     */
    public resetForm = () => {
      // Will reset the state of disable text field - call setstate to change state
      // Will clear text for active text field - 
      console.log("resetForm called!!");
      this.setState({
        customerData:{
          CustomerName:"",
          CustomerId : null,
        },
        productData: {
          ProductId: null,
          ProductName: "",
          ProductUnitPrice: "",
          ProductExpiryDate: "",
          ProductType: "",
          NumberofUnits: "",
          TotalValue: "",
        },
        status : "Reset Done!!"
      });

  }

  public render(): React.ReactElement<ISalesPnpProps> {
    return (
      <div className={styles.salesPnp}>
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.column}>
              <span className={styles.title}>Welcome to SharePoint!</span>
              {/* <p className={styles.subTitle}>
                Customize SharePoint experiences using Web Parts.
              </p> */}
            </div>
          </div>
          <hr />
          <div>
            <Pivot
              aria-label="Large Link Size Pivot Example"
              linkSize={PivotLinkSize.large}
            >
              <PivotItem headerText="Add ">
                <Label>To add Orders Fill the form below.</Label>
                <div className={styles.emptyheight}></div>
              </PivotItem>
              <PivotItem headerText="Update">
                <Label>Select Order Id below.</Label>
                {/* <TextField
                  prefix="Order Id"
                  type="number"
                  min={1}
                  required
                  value={this.state.OrderIdList}
                  onChange={this.getOrderDetailsToUpdate}
                /> */}
              <Dropdown
              required
              selectedKey={[this.state.orderId]}
              prefix="Order Id"
              options={this.state.orderIdList}
              onChange={this.getOrderDetailsToUpdate}
              ></Dropdown>

              </PivotItem>
              <PivotItem headerText="Delete">
                <Label>To update Orders Enter Order Id below.</Label>
                <div className={styles.emptyheight}></div>
              </PivotItem>
            </Pivot>
            {/* <div className={styles.emptyheight}>
              {this.state.customerData.CustomerName === "Aman" && (
                <h1>Messages</h1>
              )}
            </div> */}
            <Dropdown
              required
              selectedKey={[this.state.customerData.CustomerId]}
              id="forReset1"
              label="Enter Customer Name"
              options={this.state.customerNameList}
              onChange={this.getCustomerName}
              ></Dropdown>

            <Stack horizontal wrap tokens={stackTokens}>
              <Stack tokens={bigVertStack}>
                <Dropdown
                  required
                  selectedKey={[this.state.productData.ProductId]}
                  id="forReset2"
                  label="Enter Product Name"
                  options={this.state.productNameList}
                  onChange={this.getProductName}
                  styles={dropdownStyles}
                ></Dropdown>
                <TextField
                  id="forReset3"
                  label="Number of Units"
                  type="number"
                  min={1}
                  required
                  value={this.state.productData.NumberofUnits}
                  onChange={this.setNumberofUnits}
                />
              </Stack>
              <Stack tokens={SmallVertStack}>
                <TextField
                  label="Product Type"
                  disabled
                  placeholder={this.state.productData.ProductType}
                />
                <TextField
                  label="Product Expiry Date"
                  disabled
                  placeholder={
                    this.state.productData.ProductExpiryDate === ""
                      ? ""
                      : new Date(
                          this.state.productData.ProductExpiryDate
                        ).toDateString()
                  }
                />
                <TextField
                  label="Product Unit Price"
                  disabled
                  placeholder={this.state.productData.ProductUnitPrice}
                />
              </Stack>
            </Stack>
            <div>
              <Label>Total Sales Price</Label>
              <TextField
                ariaLabel="disabled Product Sales Price"
                readOnly
                prefix="Rs. "
                placeholder={this.state.productData.TotalValue}
              />
            </div>
            <div className={styles.emptyheight}>{this.state.status}</div>
            <div className={styles.emptyheight}></div>
            <Stack horizontal tokens={stackTokens}>
              <DefaultButton
                text="Create"
                onClick={this.validateItem}
              ></DefaultButton>
              <DefaultButton
                text="Reset"
                onClick={() => this.resetForm()}
              ></DefaultButton>

              <DefaultButton
                text="Update"
                onClick={() => this._spOps.updateItem(this.props.context, 1)}
              ></DefaultButton>
            </Stack>

            <div className={styles.emptyheight}></div>
          </div>
        </div>
      </div>
    );
  }
}
// () =>
//                   this._spOps
//                     .validateAndCreateitems(
//                       this.props.context,
//                       this.state.customerData,
//                       this.state.productData
//                     )
//                     .then((result: string) => {
//                       this.setState({ status: result });
//                     })
