import * as React from "react";
import styles from "./SalesPnp.module.scss";
import { ISalesPnpProps } from "./ISalesPnpProps";
import { ISalesPnpState } from "./ISalesPnpState";
import { escape } from "@microsoft/sp-lodash-subset";

import {
  DefaultButton,
  Dropdown,
  TextField,
  Stack,
  IStackTokens,
  IDropdownStyles,
} from "office-ui-fabric-react";
const stackTokens: IStackTokens = { childrenGap: 50 };
const bigVertStack: IStackTokens = { childrenGap: 60 };
const SmallVertStack: IStackTokens = { childrenGap: 20 };
const dropdownStyles: Partial<IDropdownStyles> = { dropdown: { width: 300 } };

import { Label } from "office-ui-fabric-react/lib/Label";
import {
  Pivot,
  PivotItem,
  PivotLinkSize,
} from "office-ui-fabric-react/lib/Pivot";

import { spOperation } from "../Services/spServices";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";

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
      status: "",
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
          console.log(results);
          let result: any = results.value[0];
          let date = new Date(result.ProductExpiryDate);

          console.log(result.ProductExpiryDate);
          console.log(date);

          this.setState({
            productData: {
              ProductName: data.text,
              ProductId: data.key,
              ProductExpiryDate: date,
              ProductType: result.ProductType,
              ProductUnitPrice: result.Product_x0020_Unit_x0020_Price,
              NumberofUnits: this.state.productData.NumberofUnits,
              TotalValue: this.state.productData.TotalValue,
            },
          });
        });
      });
  };

  /**
   * setNumberofUnits
   */
  public setNumberofUnits = (event: any, data: any) => {
    if (this.state.productData.ProductUnitPrice !== "") {
      let numberofUnits: number = parseInt(data);
      let priceofunit: number = parseInt(
        this.state.productData.ProductUnitPrice
      );
      let totalValue: number = numberofUnits * priceofunit;
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
    }
    return;
  };

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
                <div className={styles.emptyheight}></div>
                <Dropdown
                  required
                  label="Enter Customer Name"
                  options={this.state.customerNameList}
                  onChange={this.getCustomerName}
                ></Dropdown>

                <Stack horizontal tokens={stackTokens}>
                  <Stack tokens={bigVertStack}>
                    <Dropdown
                      required
                      label="Enter Product Name"
                      options={this.state.productNameList}
                      onChange={this.getProductName}
                      styles={dropdownStyles}
                    ></Dropdown>
                    <TextField
                      label="Number of Units"
                      type="number"
                      required
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
                    disabled
                    placeholder={this.state.productData.TotalValue}
                  />
                </div>
                <div className={styles.emptyheight}>{this.state.status}</div>
                <div className={styles.emptyheight}></div>
                <Stack horizontal tokens={stackTokens}>
                  <DefaultButton
                    text="Add"
                    onClick={() =>
                      this._spOps
                        .validateAndAdditems(
                          this.props.context,
                          this.state.customerData,
                          this.state.productData
                        )
                        .then((result: string) => {
                          this.setState({ status: result });
                        })
                    }
                  ></DefaultButton>

                  <DefaultButton
                    text="Delete"
                    onClick={() =>
                      this._spOps.updateItem(this.props.context, 1)
                    }
                  ></DefaultButton>
                </Stack>
                <div className={styles.emptyheight}></div>
              </PivotItem>
              <PivotItem headerText="Recent">
                <Label>Pivot #2</Label>
              </PivotItem>
              <PivotItem headerText="Shared with me">
                <Label>Pivot #3</Label>
              </PivotItem>
            </Pivot>
          </div>
        </div>
      </div>
    );
  }
}
