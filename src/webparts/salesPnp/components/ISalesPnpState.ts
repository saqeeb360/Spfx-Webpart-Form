import { IDropdownOption } from "office-ui-fabric-react";

export interface ISalesPnpState {
  customerNameList : IDropdownOption[];
  productNameList : IDropdownOption[];
  orderIdList : IDropdownOption[];
  customerData : any;
  productData : any;
  orderId : any;
  status : string;
}
