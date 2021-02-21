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
  PrimaryButton,
  IStyleSet,
  Slider,
} from "office-ui-fabric-react";
import { Label } from "office-ui-fabric-react/lib/Label";
import {
  IPivotStyles,
  Pivot,
  PivotItem,
  PivotLinkFormat,
  PivotLinkSize,
} from "office-ui-fabric-react/lib/Pivot";
import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
const pivotStyles: Partial<IStyleSet<IPivotStyles>> = {
  link: { width: "22%" },
  linkIsSelected: { width: "22%" },
};
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
      orderIdList: [],
      CustomerName: "",
      CustomerId: "",
      ProductId: "",
      ProductName: "",
      ProductUnitPrice: "",
      ProductExpiryDate: "",
      ProductType: "",
      NumberofUnits: "7",
      TotalValue: "",
      orderId: "",
      status: "",
      whichButton: "Create",
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
    this._spOps.getOrderList(this.props.context).then((result: any) => {
      this.setState({ orderIdList: result });
    });
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
      CustomerName: data.text,
      CustomerId: data.key,
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
    this._spOps
      .getProductDetails(this.props.context, data)
      .then((result: any) => {
        let date = new Date(result.ProductExpiryDate);
        var totalValue: any;
        if (this.state.NumberofUnits === "") {
          // Update Total Value when Number of Unit is not zero! otherwise don't update!!
          totalValue = this.state.TotalValue;
        } else if (this.state.NumberofUnits === "0") {
          totalValue = this.state.TotalValue;
        } else {
          totalValue =
            result.Product_x0020_Unit_x0020_Price * this.state.NumberofUnits;
        }
        // console.log(results);
        // console.log(result.ProductExpiryDate);
        // console.log(date);
        this.setState({
          ProductName: data.text,
          ProductId: data.key,
          ProductExpiryDate: date,
          ProductType: result.ProductType,
          ProductUnitPrice: result.Product_x0020_Unit_x0020_Price,
          TotalValue: totalValue,
        });
      });
  };
  /**
   * setNumberofUnits is called when number of units is changed to store the value in state.
   */
  public setNumberofUnits = (event: any, data: any) => {
    console.log("setNumberofUnits called!!");
    // console.log(this.state.ProductUnitPrice, data);
    var numberofUnits: any;
    var totalValue: any;
    if (data === "0") {
      numberofUnits = data;
      totalValue = "";
    } else if (data === "") {
      console.log(
        "setNumberofUnits called -> In ifelse -> data = '' statement!!"
      );
      numberofUnits = data;
      totalValue = "";
    } else if (this.state.ProductUnitPrice === "") {
      console.log(
        "setNumberofUnits called -> In ifelse -> UnitPrice = '' statement!!"
      );
      numberofUnits = parseInt(data);
      totalValue = "";
    } else {
      console.log("setNumberofUnits called -> In else statement!!");
      numberofUnits = parseInt(data);
      var priceofunit: number = parseInt(this.state.ProductUnitPrice);
      totalValue = numberofUnits * priceofunit;
    }
    // console.log(numberofUnits);
    // console.log(priceofunit);
    // console.log(totalValue);
    this.setState({
      NumberofUnits: numberofUnits,
      TotalValue: totalValue,
    });
    return;
  };
  /**
   * validateItemAndAdd and upload the new item
   */
  public validateItemAndAdd = () => {
    console.log("validateItemAndAdd called!!");
    let myStateList = [
      this.state.CustomerId,
      this.state.CustomerName,
      this.state.ProductId,
      this.state.ProductName,
      this.state.ProductType,
      this.state.ProductUnitPrice,
      this.state.NumberofUnits,
      this.state.TotalValue,
    ];
    console.log(myStateList);
    for (let i = 0; i < myStateList.length; i++) {
      if (myStateList[i] === "") {
        this.setState({ status: "Error! Details missing." });
        return;
      }
    }

    console.log("Validate Complete and Uploading Order Details");
    this._spOps
      .createItems(this.props.context, this.state)
      .then((result: string) => {
        this.setState({ status: result });
      });
  };

  /**
   * validateItemAndModify
   */
  public validateItemAndModify = () => {
    if (this.state.orderId === "" || this.state.orderId === null) {
      this.setState({ status: "Error! Order ID missing." });
      return;
    } else {
      this._spOps.updateItem(this.state).then((status) => {
        this.setState({ status: status });
      });
    }
  };
  /**
   * validateAndDelete
   */
  public validateAndDelete = () => {
    if (this.state.orderId === "" || this.state.orderId === null) {
      this.setState({ status: "Error! Order ID missing." });
      return;
    } else {
      this._spOps.deleteItem(this.state.orderId).then((response) => {
        this.setState({ status: response });
      });
    }
  };
  /**
   * getOrderDetailsToUpdate is called when order id field is changed to get the item details.
   */
  public getOrderDetailsToUpdate = (event: any, data: any) => {
    // Valid Order Id -> Not empty -> not zero
    console.log("getOrderDetailsToUpdate called!");
    if (data === "") {
      console.log("Empty data");
      return;
    }
    // Now get the Order List details from rest and call setstate to change the state
    this._spOps.getUpdateitem(this.props.context, data).then((results) => {
      var result = results.value[0];
      // console.log(result);
      // find customer name
      var customerId = result.Customer_x0020_IDId;
      var customerName;
      this.state.customerNameList.forEach((item) => {
        var flag = false;
        if (item.key === customerId && flag === false) {
          customerName = item.text;
          flag = true;
        }
      });

      var productId = result.Product_x0020_IDId;
      var productName;
      this.state.productNameList.forEach((item) => {
        var flag = false;
        if (item.key === productId && flag === false) {
          productName = item.text;
          flag = true;
        }
      });

      // Now find the product details to fill
      var data1 = { key: result.Product_x0020_IDId, text: productName };
      this.getProductName({}, data1);

      this.setState({
        orderId: data.key,
        CustomerName: customerName,
        CustomerId: customerId,
        NumberofUnits: result.UnitsSold,
      });
    });
  };
  /**
   * controlTabButton is called when tabs are changed to store which tab is active.
   */
  public controlTabButton = (data: any) => {
    console.log("Tab Changed");
    console.log(data);
    this.resetForm();
    // if (data.props.itemKey === "1") {
    //   // Add tab clicked
    //   // reset the tab and setstate for button
    //   this.setState({ whichButton: "Create" });
    // } else if (data.props.itemKey === "2") {
    //   this.setState({ whichButton: "Update" });
    // } else if (data.props.itemKey === "3") {
    //   this.setState({ whichButton: "Delete" });
    // }
  };
  /**
   * resetForm
   */
  public resetForm = () => {
    // Will reset the state of disable text field - call setstate to change state
    // Will clear text for active text field -
    console.log("resetForm called!!");
    this.setState({
      orderId: null,
      CustomerName: "",
      CustomerId: null,
      ProductId: null,
      ProductName: "",
      ProductUnitPrice: "",
      ProductExpiryDate: "",
      ProductType: "",
      NumberofUnits: "",
      TotalValue: "",
      status: "",
    });
    this.componentDidMount();
  };
  /**
   * renderButton is used to show active tab's button - eg: for ADD tab button should be SAVE button
   */
  // public renderButton = () => {
  //   if (this.state.whichButton === "Create") {
  //     return (
  //       <PrimaryButton
  //         text="SAVE"
  //         onClick={this.validateItemAndAdd}
  //       ></PrimaryButton>
  //     );
  //   } else if (this.state.whichButton === "Update") {
  //     return (
  //       <PrimaryButton
  //         text="MODIFY"
  //         onClick={this.validateItemAndModify}
  //       ></PrimaryButton>
  //     );
  //   } else if (this.state.whichButton === "Delete") {
  //     return (
  //       <PrimaryButton
  //         text="DELETE"
  //         onClick={this.validateAndDelete}
  //         // onClick={() =>
  //         //   this._spOps
  //         //     .deleteItem(this.props.context, this.state.orderId)
  //         //     .then((status) => {
  //         //       this.setState({ status: status });
  //         //     })
  //         // }
  //       ></PrimaryButton>
  //     );
  //   }
  // }
/**
 * productDetailBox
 */
public productDetailBox = () => {
  // If status is empty then check Product name is filled or not
  if(this.state.status === ""){
    if(this.state.ProductName === ""){
      return  (
      <Stack
        className={styles.productDetails}
        tokens={SmallVertStack}
      >
        Please Fill All Fields.
      </Stack>
    )}
    else {
      return (
      <Stack
        className={styles.productDetails}
        tokens={SmallVertStack}
      >
        <div>
          {this.state.ProductType === ""
            ? ""
            : "Type : " + this.state.ProductType}
        </div>
        <div>
          {this.state.ProductExpiryDate === ""
            ? ""
            : "Available Before : " +
              new Date(this.state.ProductExpiryDate)
                .toDateString()
                .slice(4, 7) +
              " " +
              new Date(this.state.ProductExpiryDate)
                .toDateString()
                .slice(-4)}
        </div>
        <div>
          {this.state.ProductUnitPrice === ""
            ? ""
            : "₹ " + this.state.ProductUnitPrice + "/Unit"}
        </div>
      </Stack>
    )}
  }
  else {
    return (
      <Stack
        className={styles.productDetails}
        tokens={SmallVertStack}
      >
        {this.state.status}
      </Stack>
    )
  }
}

/**
 * deleteProductBox
 */
public deleteProductBox = () => {
  if(this.state.status === ""){
    if(this.state.orderId === "" ||this.state.orderId === null ){
      return (
        <Stack
          className={styles.allProduct}
          tokens={SmallVertStack}
        >
          Please Fill Order ID.
        </Stack>
      ) 
    } 
    else{
      return (
        <Stack
        className={styles.allProduct}
        tokens={SmallVertStack}
        >
          <div>
            {"Name : " + this.state.CustomerName}
          </div>
          <div>
            {"Car : " + this.state.ProductName}
          </div>
          <div>
            {"Type : " + this.state.ProductType}
          </div>
          <div>
            {"Available Before : " +
                new Date(this.state.ProductExpiryDate)
                  .toDateString()
                  .slice(4, 7) +
                " " +
                new Date(this.state.ProductExpiryDate)
                  .toDateString()
                  .slice(-4)}
          </div>
          <div>
            {"Unit Price :  ₹ " + this.state.ProductUnitPrice + "/Unit"}
          </div>
          <div>
            {"Number of Unit :" + this.state.NumberofUnits}
          </div>
          <div>
            {"Total Price  ₹ " + this.state.TotalValue}
          </div>
        </Stack>
      )
    }
  }
  else {
    return (
    <Stack
      className={styles.allProduct}
      tokens={SmallVertStack}
    >
      {this.state.status}
    </Stack>
    )
  }
}

  public render(): React.ReactElement<ISalesPnpProps> {
    return (
      <div className={styles.salesPnp}>
        <div className={styles.row}>
          <div className={styles.column}>
            <img
              src={require("./assests/logo.png")}
              alt="logo"
              className={styles.logo}
            />
            <span className={styles.title}>Grace, space, pace</span>
          </div>
        </div>
        <hr />
        <div className={styles.container}>
          <div>
            <Pivot
              styles={pivotStyles}
              aria-label="Large Link Size Pivot Example"
              linkSize={PivotLinkSize.large}
              linkFormat={PivotLinkFormat.tabs}
              onLinkClick={this.controlTabButton}
            >
              <PivotItem headerText="ORDER NOW" itemKey="1" itemIcon="AddTo">
                {/* <Label>To add Orders Fill the form below.</Label> */}
                <div className={styles.emptyheight}></div>
                <Dropdown
                  required
                  selectedKey={[this.state.CustomerId]}
                  id="forReset1"
                  label="Customer Name"
                  options={this.state.customerNameList}
                  onChange={this.getCustomerName}
                ></Dropdown>

                <Stack horizontal wrap tokens={stackTokens}>
                  <Stack tokens={bigVertStack}>
                    <Dropdown
                      required
                      selectedKey={[this.state.ProductId]}
                      id="forReset2"
                      label="Product Name"
                      options={this.state.productNameList}
                      onChange={this.getProductName}
                      styles={dropdownStyles}
                    ></Dropdown>
                    {/* <TextField
                  id="forReset3"
                  label="Number of Units"
                  type="number"
                  min={1}
                  required
                  value={this.state.NumberofUnits}
                  onChange={this.setNumberofUnits}
                /> */}
                    <Slider
                      label="Number of Units"
                      min={0}
                      max={50}
                      step={1}
                      onChanged={this.setNumberofUnits}
                      value={this.state.NumberofUnits}
                    />
                  </Stack>
                  {this.productDetailBox()}
                  {/* <TextField
                  label="Product Type"
                  disabled
                  placeholder={this.state.ProductType}
                />
                <TextField
                  label="Product Expiry Date"
                  disabled
                  placeholder={
                    this.state.ProductExpiryDate === ""
                      ? ""
                      : new Date(this.state.ProductExpiryDate).toDateString()
                  }
                />
                <TextField
                  label="Product Unit Price"
                  disabled
                  placeholder={this.state.ProductUnitPrice}
                /> */}
                </Stack>
                <div>
                  <Label>Total Sales Price</Label>
                  <TextField
                    ariaLabel="disabled Product Sales Price"
                    readOnly
                    prefix="₹ "
                    placeholder={this.state.TotalValue}
                  />
                </div>
                <div className={styles.emptyheight}></div>
                <Stack
                  horizontal
                  tokens={stackTokens}
                  className={styles.buttonContainer}
                >
                  <PrimaryButton
                    text="SAVE"
                    className={styles.primaryButton}
                    onClick={this.validateItemAndAdd}
                  ></PrimaryButton>

                  <DefaultButton
                    className={styles.primaryButton}
                    text="CLEAR"
                    onClick={() => this.resetForm()}
                  ></DefaultButton>
                </Stack>
              </PivotItem>
              <PivotItem headerText="UPDATE" itemKey="2" itemIcon="Handwriting">
                <div className={styles.emptyheight}></div>
                <Dropdown
                  required
                  selectedKey={[this.state.orderId]}
                  label="Order Id"
                  options={this.state.orderIdList}
                  onChange={this.getOrderDetailsToUpdate}
                ></Dropdown>
                <Stack horizontal wrap tokens={stackTokens}>
                  <Stack tokens={bigVertStack}>
                    <Dropdown
                      required
                      selectedKey={[this.state.CustomerId]}
                      label="Customer Name"
                      options={this.state.customerNameList}
                      onChange={this.getCustomerName}
                      styles={dropdownStyles}
                    ></Dropdown>
                    <Dropdown
                      required
                      selectedKey={[this.state.ProductId]}
                      label="Product Name"
                      options={this.state.productNameList}
                      onChange={this.getProductName}
                      styles={dropdownStyles}
                    ></Dropdown>
                    <Slider
                      label="Number of Units"
                      min={0}
                      max={50}
                      step={1}
                      onChanged={this.setNumberofUnits}
                      value={this.state.NumberofUnits}
                    />
                  </Stack>
                  {this.productDetailBox()}
                </Stack>
                <div>
                  <Label>Total Sales Price</Label>
                  <TextField
                    ariaLabel="disabled Product Sales Price"
                    readOnly
                    prefix="₹ "
                    placeholder={this.state.TotalValue}
                  />
                </div>
                <Stack
                  horizontal
                  tokens={stackTokens}
                  className={styles.buttonContainer}
                >
                  <PrimaryButton
                    className={styles.primaryButton}
                    text="MODIFY"
                    onClick={this.validateItemAndModify}
                  ></PrimaryButton>
                  <DefaultButton
                    className={styles.primaryButton}
                    text="CLEAR"
                    onClick={() => this.resetForm()}
                  ></DefaultButton>
                </Stack>
              </PivotItem>
              <PivotItem headerText="DELETE" itemKey="3" itemIcon="Delete">
                <div className={styles.emptyheight}></div>
                <Dropdown
                  required
                  selectedKey={[this.state.orderId]}
                  label="Order Id"
                  options={this.state.orderIdList}
                  onChange={this.getOrderDetailsToUpdate}
                ></Dropdown>
                  {this.deleteProductBox()}
                  <Stack horizontal tokens={stackTokens} className={styles.buttonContainer}>
                    <PrimaryButton
                      className={styles.primaryButton}
                      text="DELETE"
                      onClick={this.validateAndDelete}
                    ></PrimaryButton>
                    <DefaultButton
                      className={styles.primaryButton}
                      text="CLEAR"
                      onClick={() => this.resetForm()}
                    ></DefaultButton>
                  </Stack>

              </PivotItem>
            </Pivot>
            <div className={styles.emptyheight}></div>
          </div>
        </div>
      </div>
    );
  }
}
