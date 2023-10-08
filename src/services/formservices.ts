import { SPFI, SPFx} from "@pnp/sp";

import { getSP } from "../pnpjsconfig";
import {  ITrc, IMolaceContractRequestFormKamSolutionProps } from "../webparts/molaceContractRequestFormKamSolution/components/IMolaceContractRequestFormKamSolutionProps";
import * as formconst from "../webparts/constant";
import { Web } from "@pnp/sp/webs";



  
  export const getLatestItemId= async (props:IMolaceContractRequestFormKamSolutionProps):Promise<ITrc[]>=> {
    const _sp :SPFI = getSP(props.context) ;
    const items = await _sp.web.lists.getByTitle(formconst.LISTNAME).items.orderBy("ID", false).top(1)();
    return items;
    //return items.length > 0 ? items[0].ID : 0;
  }

  
  
/*   export const getCustomerItems= async (props:IMolaceContractRequestFormKamSolutionProps):Promise<ICustomer[]>=> {
    //const _sp :SPFI = getSP(props.context) ;
    const _web = Web(formconst.CUSTOMER_LIST_URL).using(SPFx(props.context));
    //const _web = Web([_sp.web, "/sites/DemoTest"]);

    const items =_web.lists.getByTitle(formconst.CUSTOMER_LISTNAME).items.select().orderBy('Title',true)();
    console.log('Customer items',items);
    return items;
      
      
  } */
 
  export const getComplianceText = (props: IMolaceContractRequestFormKamSolutionProps) => {
    const _sp: SPFI = getSP(props.context);
    return _sp.web.lists.getByTitle(formconst.COMPLIANCE_LIST_NAME).items
      .select('ComplianceText')
      .orderBy('ID', true)()
      .then((items) => {
        const complianceTextArray = items.map((item) => item.ComplianceText).toString();
        //console.log(complianceTextArray)
      return complianceTextArray;
      })
      .catch((error) => {
        console.error("Error:", error);
        throw error; // Rethrow the error to propagate it to the caller
      });
  };

  export const getCustomerRef=(props:IMolaceContractRequestFormKamSolutionProps,customerName: string) => {
    //console.log(customerName)
    const _web = Web(formconst.CUSTOMER_URL).using(SPFx(props.context));
    return new Promise((resolve, reject) => {
      _web.lists.getByTitle(formconst.CUSTOMER_LISTNAME).items.select("RefCode").filter(`Title eq '${customerName}'`)()
        .then((items) => {
          if (items.length > 0) {
            const customerRef = items[0].RefCode;
            
            resolve(customerRef);
          } else {
            reject(new Error("Customer not found"));
          }
        })
        .catch((error) => {
          reject(error);
        });
    });
  } 
  
  
  
  export const submitDataAndGetId = async (props:IMolaceContractRequestFormKamSolutionProps,data:any,listFolderpath: string): Promise<any> => {
  
    const _sp :SPFI = getSP(props.context) ;
   // _sp.web.folders.addUsingPath(listFolderpath);
    return _sp.web.lists.getByTitle(formconst.LISTNAME).addValidateUpdateItemUsingPath([
      { FieldName: 'Title', FieldValue: data.Title },
      { FieldName: 'TypeofRequest', FieldValue: data.TypeofRequest },
      { FieldName: 'PaymentTerms', FieldValue: data.PaymentTerms },
      { FieldName: 'OtherConditions', FieldValue: data.OtherConditions },
      { FieldName: 'ApplicableLaw', FieldValue: data.ApplicableLaw },
      { FieldName: 'PaymentColection', FieldValue: data.PaymentColection },
      { FieldName: 'BLClause', FieldValue: data.BLClause },
      { FieldName: 'Background', FieldValue: data.Background },
      { FieldName: 'Compliance', FieldValue: data.Compliance }
    ], listFolderpath)
      .then((response) => {
        console.log(response)
        //console.log("New Item",response[8].FieldValue)
     
          // Send item ID as a response to the promise
          const itemId = response[9].FieldValue;
          console.log("ID",itemId)
          // Resolve the promise with the item ID
          return Promise.resolve(itemId);
      })
      .catch((error) => {
          // Handle any errors that occurred during the request
          return Promise.reject(error);
      });

    
  }

  export const submitMatrixData = async (props: IMolaceContractRequestFormKamSolutionProps, data: any, listFolderpath: string): Promise<any> => {
    const _sp :SPFI = getSP(props.context) ;
   console.log("data",data)
   return _sp.web.lists.getByTitle(formconst.OUTLINE_LIST_NAME).addValidateUpdateItemUsingPath([
    { FieldName: 'Title', FieldValue: data.Title },
    { FieldName: 'POL', FieldValue: data.POL },
    { FieldName: 'POD', FieldValue: data.POD },
    { FieldName: 'VolumePerYear', FieldValue: data.VolumePerYear.toString()},
    { FieldName: 'BaseFreight', FieldValue: data.BaseFreight },
    { FieldName: 'Surchages', FieldValue: data.Surchages }
  ], listFolderpath)
    .then((response) => {
      console.log(response)
      //console.log("New Item",response[8].FieldValue)
    
        // Send item ID as a response to the promise
       // const itemId = response[8].FieldValue;
        //console.log("ID",itemId)
        // Resolve the promise with the item ID
        return Promise.resolve();
    })
    .catch((error) => {
        // Handle any errors that occurred during the request
        return Promise.reject(error);
    });
     
       
   
    } 
    



  

  export const updateData=(props:IMolaceContractRequestFormKamSolutionProps ,itemId: number, data: any): Promise<void>=> {
    const _sp :SPFI = getSP(props.context) ;
    return new Promise<void>((resolve, reject) => {
      _sp.web.lists.getByTitle(formconst.LISTNAME).items.getById(itemId).update(data)
        .then(() => {
          
          //console.log(e.response.headers.get("content-length"))
          resolve();
        })
        .catch((error) => {
   
          reject(error);
        });
    });
  }


  export const getOfficeRef=(props:IMolaceContractRequestFormKamSolutionProps,officeName: string) => {
    console.log(officeName)
    const _sp :SPFI = getSP(props.context) ;
    return new Promise((resolve, reject) => {
      _sp.web.lists.getByTitle(formconst.REQUEST_LISTNAME).items.select("RefCode").filter(`Title eq '${officeName}'`)()
        .then((items) => {
          if (items.length > 0) {
            const officeRef = items[0].RefCode;
            console.log(officeRef)
            resolve(officeRef);
          } else {
            reject(new Error("Office not found"));
          }
        })
        .catch((error) => {
          reject(error);
        });
    });
  }


  export async function checklistFolderExistence(props:IMolaceContractRequestFormKamSolutionProps,folderPath: string): Promise<boolean> {
   
    folderPath = folderPath.replace(formconst.BASE_URL, "");
      const _sp :SPFI = getSP(props.context) ;
      console.log(folderPath);
      const listfolder = await  _sp.web.getFolderByServerRelativePath(folderPath).select('Exists')();
      if(listfolder.Exists){
      return true;
      }else{return false;} // Folder exists
    
  }
  export async function checklibFolderExistence(props:IMolaceContractRequestFormKamSolutionProps,folderPath: string): Promise<boolean> {
    folderPath = folderPath.replace(formconst.BASE_URL, "");
    console.log(folderPath)
    const _sp :SPFI = getSP(props.context) ;
    const listfolder = await _sp.web.getFolderByServerRelativePath(folderPath).select('Exists')();
    if(listfolder.Exists){
    return true;
    }else{return false;} // Folder exists
  
}


  

   
  

 
  
  
  
  
  
  
  
  

