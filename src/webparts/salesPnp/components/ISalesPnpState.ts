import { IDropdownOption } from "office-ui-fabric-react";

export interface ISalesPnpState {
  customerNameList : IDropdownOption[];
  productNameList : IDropdownOption[];
  customerData : any;
  productData : any;
  status : string;
}
