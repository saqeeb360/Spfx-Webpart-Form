import { IDropdownOption } from "office-ui-fabric-react";

export interface ISalesPnpState {
  customerNameList : IDropdownOption[];
  productNameList : IDropdownOption[];
  orderIdList : IDropdownOption[];
  CustomerName : any;
  CustomerId : any;
  ProductId : any;
  ProductName : any;
  ProductUnitPrice : any;
  ProductExpiryDate : any;
  ProductType : any;
  NumberofUnits : any;
  TotalValue : any;
  orderId : any;
  status : string;
  whichButton : string,
}
