import * as React from 'react';

import { IMolaceContractRequestFormKamSolutionProps } from './IMolaceContractRequestFormKamSolutionProps';
import { IMolaceContractRequestFormKamSolutionState } from './IMolaceContractRequestFormKamSolutionState';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker"; 
import styles from "./Trpreqfrm.module.scss"

//import { RichText } from "@pnp/spfx-controls-react/lib/RichText";
import { getSP } from '../../../pnpjsconfig';
import { SPFI} from '@pnp/sp';
import "@pnp/sp/files";
import "@pnp/sp/folders";
import { DateConvention, DateTimePicker } from '@pnp/spfx-controls-react/lib/DateTimePicker';
import * as moment from 'moment';


//import { ProgressIndicator, Stack } from 'office-ui-fabric-react';
import "@pnp/sp/site-users/web";
import "@pnp/sp/items";
import { TextField } from '@fluentui/react/lib/TextField';
import { PrimaryButton } from '@fluentui/react';
import { Checkbox, Dropdown, IStyleFunctionOrObject, ITextFieldStyleProps, ITextFieldStyles, Label, MessageBar, MessageBarType, Spinner, SpinnerSize, Stack, getId } from 'office-ui-fabric-react';
import { ListItemPicker} from '@pnp/spfx-controls-react';
import * as ReactDOM from 'react-dom';
import "@pnp/sp/folders";
import * as formconst from "../../constant";
import 'react-quill/dist/quill.snow.css'; 
import { checklibFolderExistence, checklistFolderExistence, getComplianceText, getCustomerRef, getOfficeRef, submitDataAndGetId, submitMatrixData, updateData } from '../../../services/formservices';
//import { isEmpty } from '@microsoft/sp-lodash-subset';
import DynamicTable from './insert';
import { isEmpty } from '@microsoft/sp-lodash-subset';


  let listId: number;
  let customerreference:string;
  let officereference:string;
  let isselectedApplicant:boolean = true ;

  let isemailInvalid:boolean = false;

  let isbuttondisbled : boolean = false;
  let buttontext : string = "Submit";
  let listFolderpath:string;
  let outlinelistFolderpath:string;
  let libfolderUrl:string;
  let isLoading : boolean = false;
  let ispaymenttermcollectvisible : boolean = false;
  let ispaymenttermyesvisible : boolean = false;
  let isblothervisible : boolean = false;
  //let compstr:string
 // let cweburl:any;
  //let contactlist :string;
 

  const textFieldStyles: Partial<ITextFieldStyles> = {
    field: {
      width: '500px', // Adjust the desired width
    },
  };
  
  
  

export default class Trpreqfrm extends React.Component<IMolaceContractRequestFormKamSolutionProps, IMolaceContractRequestFormKamSolutionState> {

  private bdt: DataTransfer; 
  private vdt: DataTransfer; 
  private odt: DataTransfer; 
  
 
  filesNamesRef: React.RefObject<HTMLSpanElement>;
  handleChange: any;
  pickerId: string;
  
  constructor(props: IMolaceContractRequestFormKamSolutionProps, state: IMolaceContractRequestFormKamSolutionState) {  
    super(props);  
    this.bdt = new DataTransfer();
    this.vdt = new DataTransfer();
    this.odt = new DataTransfer();
    this.filesNamesRef = React.createRef();
    this.pickerId = getId('inline-picker');
    this.state = {  
      title: "",  
      users: [], 
      usersstr:"",
      partyusers: [],
      ApplicantId:0,
      ValueDropdown:"",
      customerlist:"",
      startDate:new Date(),
      endDate:new Date(),
      dateduration:"0",
      cargodescription:"",
      contractval:"",
      iscontractvalValid: true,
      portpairs:"",
      freight:"",
      othercon:"",
      applaw:"",
     showProgress:false,
      progressLabel: "File upload progress",
      progressDescription: "",
      progressPercent: 0,
      voyage:"",
      background:"",
      addothers:"",
      InterestedPartiesId:0,
      isSuccess: false,
      files: [],
      bgdocuments:"",
      vdocuments:"",
      odocuments:"",
      interestedPartiesexternal: [],
      interestedPartiesexternalstr:"",
      newParty: "",
      selectedTags: [],
      baf:"",
      onload:true,
      allfieldsvalid:true,
      listfolderExists: true,
      libfolderExists : true,
      items: [],
      paymentterm:undefined,
      ptyes:"",
      blothers:"",
      typeofreq:undefined,
      paymenttermcollect:undefined,
      paymentcollection:undefined,
      blclause:undefined,
      comp:"No",
      compstr:"",
      customer:"",
      contactlistid:"",
      weburl:"",
      
     
      
    }; 
    
  }

  public componentDidMount()
{

  let email=this.props.userDisplayName;
  const _sp :SPFI = getSP(this.props.context ) ;
  (_sp.web.siteUsers.getByEmail(email)()).then(user=> {this.setState({ApplicantId:user.Id})});
 
 
  this.fetchComplianceText();

}

fetchComplianceText=()=>{

  getComplianceText(this.props).then((comtext: any)=>{
    this.setState({compstr:comtext})
    //compstr = comtext;
      //console.log("in fu", compstr)
      
      })
  
  
 
}




  public _getPeoplePickerItems=(items: any[]) =>{  
  
    if(items.length> 0){
      let userid =items[0].id
      this.setState({ ApplicantId: userid });
      isselectedApplicant  = true;
      //console.log('Items new:', userid );
    }else{
      this.setState({ ApplicantId: "" });
      isselectedApplicant  = false;

    }
    
  } 
 
  public _getPartiesPeoplePickerItems=(items: any[]) =>{  
   console.log(items)
    
      let selectedUsers: string[] = [];
      items.map((item) => {
        selectedUsers.push(item.id);
       
      });
       this.setState({users: selectedUsers});
      console.log('users:',selectedUsers)  
      
    } 
 
  private _oncustomerSelectedItem=async (data: { key: string; name: string }[])=> {

   if(data.length == 0){
    this.setState({customerlist:""})
  }else{
    this.setState({customerlist:data[0].name as string})
    getCustomerRef(this.props,data[0].name).then((customerRef: string)=>{

      customerreference = customerRef
      //console.log(customerRef);
      
    });

    listFolderpath = formconst.WEB_URL+"/Lists/"+ formconst.LISTNAME+"/" + data[0].name; 
    //let relistFolderpath = formconst.LISTNAME+"/" + data[0].name; 
    const listfolderExists = await checklistFolderExistence(this.props,listFolderpath);
    
   this.setState({ listfolderExists });
    console.log("List",listfolderExists);
    libfolderUrl = formconst.LIBRARYNAME +"/" + data[0].name;
    const libfolderExists = await checklibFolderExistence(this.props,libfolderUrl);
    this.setState({ libfolderExists });
    console.log("Folder",libfolderExists);

    outlinelistFolderpath = formconst.WEB_URL+"/Lists/"+ formconst.OUTLINE_LIST_NAME +"/" + data[0].name; 
    //outlinelistFolderpath = formconst.OUTLINE_LIST_NAME +"/" + data[0].name; 
    //let relistFolderpath = formconst.LISTNAME+"/" + data[0].name; 
    const outlinelistfolderExists = await checklistFolderExistence(this.props,outlinelistFolderpath);
    
   this.setState({ listfolderExists });
    console.log("List",outlinelistfolderExists);
   }

     //isselectedCustomer = data.length >0 ? true : false;
  
  }

   

  
  private _onofficeSelectedItem=(data: { key: string; name: string }[])=> {
    
    if(data.length == 0 ){
      this.setState({ValueDropdown:""})
    }else{
    this.setState({ValueDropdown:data[0].name as string})
    getOfficeRef(this.props,data[0].name).then((officeRef: string)=>{

      officereference = officeRef
      console.log(officereference);
      
    })
    }

  }

  private _onpcollectionSelectedItem=(data: { key: string; name: string }[])=> {
    
    if(data.length == 0 ){
      this.setState({paymentcollection:""})
    }else{
    this.setState({paymentcollection:data[0].name as string})
    }

  }

  private _onchangedStartDate=(stdate: any): void =>{  
    this.setState({ startDate: stdate }); 

    
    const startDate = moment(stdate);
    const timeEnd = moment(this.state.endDate);
    const diff = timeEnd.diff(startDate,'days').toString();
    //const diffDuration = moment.duration(diff)
    console.log('diffdur', diff)
    this.setState({ dateduration: diff }); 
   
  }
  private _onchangedEndDate=(eddate: any): void=> {  
    this.setState({ endDate: eddate });  
    const startDate = moment(this.state.startDate);
    const timeEnd = moment(eddate);
    const diff = (timeEnd.diff(startDate,'days')+1).toString();
    //const diffDuration = moment.duration(diff);
    console.log('diffdur', diff)
    this.setState({ dateduration: diff }); 
  }


  


 
private _newparty=(ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newText: string): void =>{ 
  this.setState({newParty:newText})
}
handleAddParty = () => {
  const { newParty, interestedPartiesexternal } = this.state;
  if (newParty.trim() !== ''&& /^\w+([.-]?\w+)*@\w+([.-]?\w+)*(\.\w{2,3})+$/.test(newParty)) {
    const updatedParties = [...interestedPartiesexternal, newParty]
    console.log(updatedParties)

    this.setState({ interestedPartiesexternal: updatedParties, newParty: '', interestedPartiesexternalstr:(JSON.stringify(updatedParties)).slice(1, -1).replace(/"/g, '')});
    isemailInvalid = false;
  } else{

    isemailInvalid = true;
    this.setState({newParty:"",allfieldsvalid:false})

  }
  //console.log(interestedPartiesexternal.toString())
};


  bghandleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
  
  const filesNames = document.querySelector<HTMLSpanElement>('#bgfilesList > #bgfiles-names');
   // console.log(files.length);
    for (let i = 0; i < e.target.files.length; i++) {
      let fileBloc = (
        <span key={i} className="file-block">
          <span className="name">{e.target.files.item(i).name}</span>
  <span className="file-delete">
     <button> Remove</button>
  </span>
  <br/>
        </span>
      );
    
      if (filesNames) {
        const fileBlocContainer = document.createElement('div');
        ReactDOM.render(fileBloc, fileBlocContainer);
        filesNames?.appendChild(fileBlocContainer.firstChild);
      //filesNames?.appendChild(fileBloc);
      }
    }
  
    for (let file of e.target.files as any) {
      this.bdt.items.add(file);
    }
  
    e.target.files = this.bdt.files;
  
    document.querySelectorAll('span.file-delete').forEach((button) => {
      button.addEventListener('click', () => {
        let name = button.nextSibling.textContent;
  
        (button.parentNode as HTMLElement)?.remove();
  
        for (let i = 0; i < this.bdt.items.length; i++) {
          if (name === this.bdt.items[i].getAsFile()?.name) {
            this.bdt.items.remove(i);
            continue;
          }
        }
  
        e.target.files = this.bdt.files;
  
      });
    });
  };

  vhandleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    
  const filesNames = document.querySelector<HTMLSpanElement>('#vfilesList > #vfiles-names');
   // console.log(files.length);
    for (let i = 0; i < e.target.files.length; i++) {
      let fileBloc = (
        <span key={i} className="file-block">
        <span className="name">{e.target.files.item(i).name}</span>
  <span className="file-delete">
    <button> Remove</button>
  </span>
  <br/>
        </span>
      );
     
      if (filesNames) {
        const fileBlocContainer = document.createElement('div');
        ReactDOM.render(fileBloc, fileBlocContainer);
        filesNames?.appendChild(fileBlocContainer.firstChild);
      //filesNames?.appendChild(fileBloc);
      }
    }
  
    for (let file of e.target.files as any) {
      this.vdt.items.add(file);
    }
  
    e.target.files = this.vdt.files;
  
    document.querySelectorAll('span.file-delete').forEach((button) => {
      button.addEventListener('click', () => {
        let name = button.nextSibling.textContent;
  
        (button.parentNode as HTMLElement)?.remove();
  
        for (let i = 0; i < this.vdt.items.length; i++) {
          if (name === this.vdt.items[i].getAsFile()?.name) {
            this.vdt.items.remove(i);
            continue;
          }
        }
  
        e.target.files = this.vdt.files;
    
      });
    });

  };
  
  ohandleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    
  const filesNames = document.querySelector<HTMLSpanElement>('#ofilesList > #ofiles-names');
   // console.log(files.length);
    for (let i = 0; i < e.target.files.length; i++) {
      let fileBloc = (
        <span key={i} className={'file-block'}>
         <span className="name">{e.target.files.item(i).name}</span>
          <span className="file-delete">
          <button> Remove</button>
         </span>
           <br/>
        </span>
      );
     
      if (filesNames) {
        const fileBlocContainer = document.createElement('div');
        ReactDOM.render(fileBloc, fileBlocContainer);
        filesNames?.appendChild(fileBlocContainer.firstChild);
      //filesNames?.appendChild(fileBloc);
      }
    }
  
    for (let file of e.target.files as any) {
      this.odt.items.add(file);
    }
  
    e.target.files = this.odt.files;
  
    document.querySelectorAll('span.file-delete').forEach((button) => {
      button.addEventListener('click', () => {
        let name = button.nextSibling.textContent;
  
        (button.parentNode as HTMLElement)?.remove();
  
        for (let i = 0; i < this.odt.items.length; i++) {
          if (name === this.odt.items[i].getAsFile()?.name) {
            this.odt.items.remove(i);
            continue;
          }
        }
  
        e.target.files = this.odt.files;
      });
    });
  };
    private _createItem  =async (props:IMolaceContractRequestFormKamSolutionProps):Promise<void>=>{

     // const _sp :SPFI = getSP(this.props.context ) ;
     if(this.state.typeofreq == undefined || !isselectedApplicant || (this.state.customerlist).length == 0 || (this.state.ValueDropdown).length == 0 ||(this.state.items).length==0
      || this.state.paymentterm == undefined || (ispaymenttermcollectvisible && this.state.paymenttermcollect == undefined) || (ispaymenttermyesvisible && isEmpty(this.state.ptyes))
       ||this.state.paymentcollection == undefined || this.state.blclause == undefined|| (isblothervisible && isEmpty(this.state.blothers)) ||isEmpty(this.state.background) || 
       isEmpty(this.state.othercon) ||this.state.comp =='No'  ||isemailInvalid||!this.state.libfolderExists ||!this.state.listfolderExists)
     {
     this.setState({allfieldsvalid:false}) ; 
     console.log(this.state.allfieldsvalid)
    
     return;
     }else {
      this.setState({allfieldsvalid:true}) ; 
      isbuttondisbled = true;
      buttontext = "Saving...";
      isLoading = true;
     }

     const tv = document.getElementById("totalvolumeID").textContent
     console.log(tv)
   
      let folderUrl: string;
      let outlinefolderUrl: string;

      
      const data = {
        Title: 'New Item creation in process',
        TypeofRequest:this.state.typeofreq,
        PaymentTerms:this.state.paymentterm,
        PaymentColection:this.state.paymentcollection,
        OtherConditions: this.state.othercon,
        ApplicableLaw: this.state.applaw,
        BLClause:this.state.blclause,
        Background: this.state.background,
        Compliance:this.state.comp
     };
      submitDataAndGetId(this.props,data,listFolderpath).then(async (itemId: any) => {
        listId = itemId   
        console.log(`Item created with ID: ${itemId}`);

        //Request ID format
        let now = new Date();
        let options: Intl.DateTimeFormatOptions = {
        day: '2-digit',
        month: '2-digit',
        year: 'numeric'
      };
      let listIdstr
       if(listId < 1000 && listId > 100){
        listIdstr = "0"+String(listId)
      }else if(listId < 100){
        listIdstr ="00"+String(listId)
      } else if(listId < 10) {
        listIdstr ="000"+String(listId)
      }else{
        listIdstr = String(listId)
      }
      
      console.log(listIdstr)
      let formattedDate = now.toLocaleDateString('en-GB', options).replace(/\//g, '');;
      let lastitemid = (listIdstr)+"-"+customerreference+"-"+officereference +"-" +formattedDate.toString();

     
     // console.log(lastitemid)
    
    //folderUrl =formconst.LIBRARYNAME + "/" + lastitemid    
   folderUrl = libfolderUrl + "/" + lastitemid
   outlinefolderUrl = outlinelistFolderpath+ "/" + lastitemid
    this.setState({title:lastitemid})
   
    console.log(outlinefolderUrl)
    
   
  }).then(async () => {
 
    await upload()
    // Update the item
    const updatedData = {
      Title: this.state.title,
      ApplicantId: this.state.ApplicantId,
      RequestingOffice: this.state.ValueDropdown,
      Customer: this.state.customerlist,
      ContractPeriodStart: this.state.startDate,
      ContractPeriodEnd: this.state.endDate,
      ContractDuration: this.state.dateduration,
      ContractVolumePerYear: tv,
      InterestedPartiesId: this.state.users,
      BackgroundSupportingDocuments: this.state.bgdocuments,
      DraftContractDocuments:this.state.vdocuments,
      InterestedPartiesExt:this.state.interestedPartiesexternalstr,
      OtherConditionsDocuments:this.state.odocuments,
      PaymentTermsCollect:this.state.paymenttermcollect,
      PaymentTermsCreditTerms:this.state.ptyes,
      BLClauseOthers:this.state.blothers,
      ComplianceAgreement:this.state.compstr
      
    };
    return updateData(this.props,listId, updatedData);
  })
   .then(() => {
    //console.log('Item Updated successfully');
    // Perform any further actions if needed
     const { items } = this.state;
  for (const itemData of items) {
  // Convert the items to a JSON object
  const jsonData = JSON.stringify(items);
  console.log(jsonData)

  const data = {
    Title: this.state.title, // Replace with a suitable title for each list item
    //jsondata: jsonData, // Store JSON data in the specified field
    POL: itemData.pol, // Replace with the appropriate field name for POL
    POD: itemData.pod, // Replace with the appropriate field name for POD
    VolumePerYear: itemData.volume, // Replace with the appropriate field name for VolumePerYear
    BaseFreight: itemData.baseFreight, // Replace with the appropriate field name for Base Freight
    Surchages: itemData.surcharges, // Replace with the appropriate field name for Surchages
  };
 
 submitMatrixData(this.props,data,outlinelistFolderpath).then(async () => {
  isbuttondisbled = false;
  buttontext = "Submit";
  isLoading = false;
  this.setState({ isSuccess: true });
  
  window.open(formconst.SUBMIT_REDIRECT,"_self")
 
})
}}) 
  .catch((error: any) => {
    
    var obj = JSON.stringify(error);
  
    if(obj.indexOf("400") !== -1)
    {    console.log("mATCH FOUND")
         // streamerror = true;
         this.setState({allfieldsvalid:false}) 
  }

    else{

    console.log('Error:', error);}
  });


const upload = async () => {

    console.log(folderUrl)
    const _sp :SPFI = getSP(props.context) ;
    let strbgurl = "";
    let vstrbgurl = "";
    let ostrbgurl = "";
    _sp.web.folders.addUsingPath(folderUrl);
    
    // bgfiles
    
    let bgfileurl = [];
    const bgcategory = 'Background'
    let bginput = document.getElementById("bgattachment") as HTMLInputElement;

    console.log(bginput.files);
  
    if (bginput.files.length > 0) {
      let bgfiles = bginput.files;
    
      for (var i = 0; i < bgfiles.length; i++) {
        let bgfile = bginput.files[i];
        console.log("bgfile",bgfile)
        bgfileurl.push(formconst.WEB_URL + "/" + folderUrl + "/" +bgfile.name);
        //console.log()
        try {
          let bguploadedFile = await _sp.web.getFolderByServerRelativePath(folderUrl).files.addChunked(bgfile.name, bgfile, (data) => {
            console.log("File uploaded successfully");
          });
          let item = await bguploadedFile.file.getItem();
          await item.update({Section:bgcategory});
          await item.update({RequestID:this.state.title})

            
        } catch (err) {
          console.error("Error uploading file:", err);
        }
      }
      let convertedStr = bgfileurl.map(url => `<a href="${url.trim()}" target="_blank">${url.substring(url.lastIndexOf("/") + 1)}</a>`);
       strbgurl = convertedStr.toString();
        //console.log(strbgurl);
        this.setState({ bgdocuments: strbgurl });
    }
      
     else {
      console.log("No file selected for upload.");
    }
    // vfiles
    let vfileurl = [];
    let vinput = document.getElementById("vattachment") as HTMLInputElement;
    const vcategory = 'Tender & draft'
    console.log(vinput.files);
    if (vinput.files.length > 0) {
      let vfiles = vinput.files;
    
      for (var i = 0; i < vfiles.length; i++) {
        let vfile = vinput.files[i];
        console.log("vfile",vfile)
        vfileurl.push(formconst.WEB_URL + "/" + folderUrl + "/" + vfile.name);
        try {
          let vuploadedFile = await _sp.web.getFolderByServerRelativePath(folderUrl).files.addChunked(vfile.name, vfile, (data) => {
            console.log("File uploaded successfully");
          });
          let item = await vuploadedFile.file.getItem();
          await item.update({Section:vcategory});
          await item.update({RequestID:this.state.title})
        } catch (err) {
          console.error("Error uploading file:", err);
        }
      }
      let vconvertedStr = vfileurl.map(url => `<a href="${url.trim()}" target="_blank">${url.substring(url.lastIndexOf("/") + 1)}</a>`);
     vstrbgurl = vconvertedStr.toString();
    //console.log(vstrbgurl);
    this.setState({ vdocuments: vstrbgurl });
    
    } else {
      console.log("No file selected for upload.");
      
    }
    
  
    // ofiles
    let ofileurl = [];
    let oinput = document.getElementById("othersattachment") as HTMLInputElement;
    const ocategory = 'Others'
    console.log(oinput.files);
   
    if (oinput.files.length > 0) {
      let ofiles = oinput.files;
   
      for (var i = 0; i < ofiles.length; i++) {
        let ofile = oinput.files[i];
        console.log("ofile",ofile)
        ofileurl.push(formconst.WEB_URL + "/" + folderUrl + "/" + ofile.name);
        try {
          let ouploadedFile = await _sp.web.getFolderByServerRelativePath(folderUrl).files.addChunked(ofile.name, ofile, (data) => {
            console.log("File uploaded successfully");
          });
          let item = await ouploadedFile.file.getItem();
          await item.update({Section:ocategory});
          await item.update({RequestID:this.state.title}) ;
             
          
        } catch (err) {
          console.error("Error uploading file:", err);
        }
      }
      let oconvertedStr = ofileurl.map(url => `<a href="${url.trim()}" target="_blank">${url.substring(url.lastIndexOf("/") + 1)}</a>`);
       ostrbgurl = oconvertedStr.toString();
      //console.log(ostrbgurl);
      this.setState({ odocuments: ostrbgurl });
      
    } else {
      console.log("No file selected for upload.");
      
    }

  }

} 

/* getTextFromItem = (item: { name: any; }) => item.name;

handleTagChange = (selectedTags: any) => {
  this.setState({ selectedTags });
};

_resetrichtext = () =>{
 
  this.setState({cargodescription:"", portpairs:"",othercon:"", applaw:"", voyage:"",background:"", addothers:"",allfieldsvalid:true})
  //streamerror = false;
  isbuttondisbled = false;
  buttontext = "Submit"

}
 */
cancel =()=>{
  //navigate(formconst.CANCEL_REDIRECT);
  window.open(formconst.CANCEL_REDIRECT,"_self");
}


handleDropdownChange = (event: React.FormEvent<HTMLDivElement>, option: { key: any; },dropdownName:string) => {
  
  console.log("field",dropdownName)
  console.log("key",option.key)
  if (dropdownName === 'typeofreq') {
    console.log("type")
    this.setState({ typeofreq: option.key });
  } else if (dropdownName === 'paymentcollection') {
    this.setState({ paymentcollection: option.key });
  } else if (dropdownName === 'blclause') {
    if(option.key === 'Others'){isblothervisible=true}else{isblothervisible=false; this.setState({blothers:""})}
    this.setState({ blclause: option.key });
  }else if (dropdownName === 'paymentterm') {
    
    if(option.key === 'Collect'){ispaymenttermcollectvisible=true}else{ispaymenttermcollectvisible=false; this.setState({paymenttermcollect:undefined})}
    this.setState({ paymentterm: option.key });
  }else if(dropdownName ==='paymenttermcollect'){
    if(option.key === 'Yes'){ispaymenttermyesvisible=true}else{ispaymenttermyesvisible=false ; this.setState({ptyes:""})}
    this.setState({ paymenttermcollect: option.key });
  }
};

handleInputChange = (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, field: string) => {
  //const { value } = (event.target as any).value;
  const { value } = event.target as HTMLInputElement;
  // Update the state based on the field parameter
  if (field === 'applaw') {
    this.setState({ applaw: value });
  } else if (field === 'background') {
    this.setState({ background: value });
  } else if (field === 'othercon') {
    this.setState({ othercon: value });
  } else if (field === 'ptyes') {
    this.setState({ ptyes: value });
  } else if (field === 'blothers') {
    this.setState({ blothers: value });
  }
}

_onCheckboxChange = (ev?: React.FormEvent<HTMLElement | HTMLInputElement>, isChecked?: boolean)=> {
  console.log(`The option has been changed to ${isChecked}.`);
  if(isChecked){
  this.setState({comp:"Yes"})
  }else{this.setState({comp:"No"})}
}


  public render(): React.ReactElement<IMolaceContractRequestFormKamSolutionProps> {
  
    //let compstrspan =(<span>compstr</span>)
    let curruser:any = this.props.userDisplayName;
    const {interestedPartiesexternal,paymentterm,typeofreq,paymenttermcollect,blclause} = this.state;
    let successMessage : JSX.Element | null
    //let textFieldErrorMessage: JSX.Element | null 
    let applicantFieldErrorMessage: JSX.Element | null 
    let customerFieldErrorMessage: JSX.Element | null 
    let OfficeFieldErrorMessage: JSX.Element | null
    //let volumeFieldErrorMessage: JSX.Element | null
   
  
    let EmailFieldErrorMessage: JSX.Element | null
    let FormFieldErrorMessage: JSX.Element | null
    let foldercustomerFieldErrorMessage: JSX.Element | null
    let typeRequestErrorMessage:JSX.Element | null
    let paymenttermErrorMessage:JSX.Element | null
    let paymentcollectionErrorMessage:JSX.Element | null
    let blclauseErrorMessage:JSX.Element | null
    let paymenttermcollectErrorMessage:JSX.Element | null
    let paymenttermyesErrorMessage:JSX.Element | null
    let blotherErrorMessage:JSX.Element | null
    let backgroundErrorMessage:JSX.Element | null
    let otherconErrorMessage:JSX.Element | null
    let complianceErrorMessage:JSX.Element | null
    let outlineFieldErrorMessage


    successMessage = this.state.isSuccess ?
    <MessageBar messageBarType={MessageBarType.success}>Request Id : {this.state.title} submitted successfully.</MessageBar>
    : null;
    
    console.log("bg",this.state.background)

    if(!this.state.allfieldsvalid){
   
      
        typeRequestErrorMessage = this.state.typeofreq == undefined  ? 
        <MessageBar messageBarType={MessageBarType.error}>Type of Request is required.</MessageBar>
        : null;

        applicantFieldErrorMessage = !isselectedApplicant ?
        <MessageBar messageBarType={MessageBarType.error}>Applicant is required.</MessageBar>
        : null;
  
        customerFieldErrorMessage = (this.state.customerlist).length == 0 ?
        <MessageBar messageBarType={MessageBarType.error}>Customer is required.</MessageBar>
        : null;

        outlineFieldErrorMessage = (this.state.items).length == 0 ?
        <MessageBar messageBarType={MessageBarType.error}>Atleast one record is required.</MessageBar>
        : null;

        paymenttermErrorMessage = this.state.paymentterm == undefined  ? 
        <MessageBar messageBarType={MessageBarType.error}>Payment Terms is required.</MessageBar>
        : null;

        paymenttermcollectErrorMessage = (ispaymenttermcollectvisible && this.state.paymenttermcollect == undefined)  ? 
        <MessageBar messageBarType={MessageBarType.error}>Payment Terms (in case of 'collect') is required.</MessageBar>
        : null;
       
        paymenttermyesErrorMessage = (ispaymenttermyesvisible && isEmpty(this.state.ptyes))  ? 
        <MessageBar messageBarType={MessageBarType.error}>Payment Terms (in case of 'Yes') is required.</MessageBar>
        : null;

        paymentcollectionErrorMessage = this.state.paymentcollection == undefined  ? 
        <MessageBar messageBarType={MessageBarType.error}>Payment Collection is required.</MessageBar>
        : null;

        blclauseErrorMessage = this.state.blclause == undefined  ? 
        <MessageBar messageBarType={MessageBarType.error}>B/L Clause is required.</MessageBar>
        : null;

        blotherErrorMessage = (isblothervisible && isEmpty(this.state.blothers))  ? 
        <MessageBar messageBarType={MessageBarType.error}>B/L Clause (Other , please specify) is required.</MessageBar>
        : null;

        backgroundErrorMessage = isEmpty(this.state.background)  ? 
        <MessageBar messageBarType={MessageBarType.error}>Background is required.</MessageBar>
        : null;
        
        otherconErrorMessage = isEmpty(this.state.othercon)  ? 
        <MessageBar messageBarType={MessageBarType.error}>Others is required.</MessageBar>
        : null;

        complianceErrorMessage = this.state.comp =='No'  ? 
        <MessageBar messageBarType={MessageBarType.error}>Compliance is required.</MessageBar>
        : null;

        foldercustomerFieldErrorMessage = (!this.state.libfolderExists || !this.state.listfolderExists) ?
        <MessageBar messageBarType={MessageBarType.error}>Kindly contact ypur IT Iteam. Folder do not exist.</MessageBar>
        : null;
  
        OfficeFieldErrorMessage = (this.state.ValueDropdown).length == 0  ?
        <MessageBar messageBarType={MessageBarType.error}>Requesting Office is required.</MessageBar>
        : null;
  
         EmailFieldErrorMessage= isemailInvalid ?
        <MessageBar messageBarType={MessageBarType.error}>Please enter a valid email address.</MessageBar>
        : null;

       /* imageFieldErrorMessage= 
        <MessageBar messageBarType={MessageBarType.error}>Stream size exceeds the allowed limit. Note the image size should be less than 80 KB</MessageBar> */
        
       /*  imageFieldErrorMessage = streamerror ? <MessageBar messageBarType={MessageBarType.blocked} isMultiline={false} onDismiss={this._resetrichtext} dismissButtonAriaLabel="Close"
        truncated={true} overflowButtonAriaLabel="See more">Stream size exceeds the allowed limit. Note that the image size in the rich text field should be less than 80 KB .
        On closing the dialog will reset the rich text field values </MessageBar>: null; */
        
        FormFieldErrorMessage= 
        <MessageBar messageBarType={MessageBarType.error}>Please provide all required information and submit the form.</MessageBar>
        
      }
   

    return (
    
    <section>
        <div>
          <p className={styles.heading}>Outline of the Agreement</p>
          <div>
          <p><span className={styles.required}><b>*</b></span> Required.</p>
    </div>
    <p className={styles.formlabel}>Type Of Request<span className={styles.required}> *</span></p>

      <Dropdown
          selectedKey={typeofreq}
          placeholder='Select Type of Request'
          options={[
            //{ key: undefined, text: 'Select' },
            { key: 'Tranportation Contract', text: 'Tranportation Contract' },
            { key: 'Bid & Tender (RFI)', text: 'Bid & Tender (RFI)' },
            { key: 'Bid & Tender (RFQ)', text: 'Bid & Tender (RFQ)' },
          
            // Add more options as needed
          ]}
          onChange={(event, option) => this.handleDropdownChange(event, option,'typeofreq')}
        /> {typeRequestErrorMessage}
        <PeoplePicker
            context={this.props.context as any}
            titleText="Applicant"
            placeholder='Select Applicant'
            defaultSelectedUsers = {[curruser]}
            personSelectionLimit={1}
            groupName={""} // Leave this blank in case you want to filter from all users
            ensureUser={true}
            showtooltip={false}
            suggestionsLimit={5}
            required={true}
            disabled={false}
            onChange={this._getPeoplePickerItems}
            showHiddenInUI={false}
            principalTypes={[PrincipalType.User]}
            resolveDelay={1000}
    /> {applicantFieldErrorMessage}

    <p className={styles.formlabel}>Requesting Office<span className={styles.required}> *</span></p>
    <ListItemPicker listId={formconst.REPORTING_OFFICE_LIST_ID}
       context={this.props.context as any}
          columnInternalName='Title'
          keyColumnInternalName='Id'
          placeholder="Select or Search office"
          substringSearch={true}
          //label="Requesting Office"
          orderBy={"Id desc"}
          itemLimit={1}
          enableDefaultSuggestions={true}
          onSelectedItem={this._onofficeSelectedItem}
          noResultsFoundText="No office Found"
          defaultSelectedItems = {[]}
    />{OfficeFieldErrorMessage}
    
  <p className={styles.formlabel}>Customer<span className={styles.required}> *</span></p>
  <ListItemPicker
                 listId={formconst.CUSTOMER_LIST_ID}
                 context={this.props.context as any}
                 webUrl={formconst.CUSTOMER_URL}
                 columnInternalName="Title"
                 keyColumnInternalName="Id"
                 placeholder="Select or Search Customer"
                 substringSearch={true}
                 orderBy={"Title"}
                 itemLimit={1}
                 enableDefaultSuggestions={true}
                 onSelectedItem={this._oncustomerSelectedItem}
                 noResultsFoundText="No Customer Found"
                 defaultSelectedItems={[]}
                />{customerFieldErrorMessage}{foldercustomerFieldErrorMessage} 
      <div>
        <table>
          <tr ><td className={styles.tabltr}>
      <p className={styles.formlabel}>Contract Period From<span className={styles.required}> *</span></p>  
      <DateTimePicker 
      //label="Contract From"
      maxDate={this.state.endDate}
            dateConvention={DateConvention.Date}
            value={this.state.startDate}  
            onChange={this._onchangedStartDate} 
            allowTextInput = {false}
            showLabels = {false}
            />
        </td>
        <td width={'100px;'}></td>
        <td className={styles.tabltr}>        
     <p className={styles.formlabel}>Contract Period To<span className={styles.required}> *</span></p>  
        
    <DateTimePicker 
    //label="Contract To"
    minDate={this.state.startDate}
          dateConvention={DateConvention.Date}
          value={this.state.endDate}  
          onChange={this._onchangedEndDate}
          allowTextInput = {false}  
          showLabels = {false}/>
          </td>
          </tr>
      </table>
    </div>  

      {/* <TextField label="Contract Duration" value={this.state.dateduration} onChange={this._onchangedduration}/> */}
       <p className={styles.formlabel}>Contract Duration (Days)</p>
      <Label>{this.state.dateduration}</Label>

      <p className={styles.formlabel}>Outline(port pairs/volume/freight)<span className={styles.required}> *</span></p>
      <DynamicTable items={this.state.items} onUpdateItems={(items) => this.setState({ items })} 
      />{outlineFieldErrorMessage}

      

      <p className={styles.formlabel}>Payment Terms<span className={styles.required}> *</span></p>

      <Dropdown
          selectedKey={paymentterm}
          placeholder='Select Payment Terms'
          options={[
            { key: 'Prepaid', text: 'Prepaid' },
            { key: 'Collect', text: 'Collect' },
          
            // Add more options as needed
          ]}
          onChange={(event, option) => this.handleDropdownChange(event, option,'paymentterm')}
        />{paymenttermErrorMessage}
{ispaymenttermcollectvisible && <div>
<p className={styles.formlabel}>Payment Terms (in case of 'collect')<span className={styles.required}> *</span></p>

<Dropdown
    selectedKey={paymenttermcollect}
    options={[
      { key: 'Yes', text: 'Yes' },
      { key: 'No', text: 'No' },
    
      // Add more options as needed
    ]}
    onChange={(event, option) => this.handleDropdownChange(event, option,'paymenttermcollect')}
  />{paymenttermcollectErrorMessage}
</div>}
{ispaymenttermyesvisible && <div>
<p className={styles.formlabel}>Payment Terms (in case of 'Yes')<span className={styles.required}> *</span></p>

<TextField value={this.state.ptyes} onChange={(event) => this.handleInputChange(event,'ptyes')} />
</div>}{paymenttermyesErrorMessage}
<p className={styles.formlabel}>Payment Collection<span className={styles.required}> *</span></p>

<ListItemPicker listId={formconst.FREIGHT_LIST_ID}
       context={this.props.context as any}
          columnInternalName='Title'
          keyColumnInternalName='Id'
          placeholder="Select your office"
          substringSearch={true}
          //label="Requesting Office"
          orderBy={"Id desc"}
          itemLimit={1}
          enableDefaultSuggestions={true}
          onSelectedItem={this._onpcollectionSelectedItem}
          noResultsFoundText="No office Found"
          defaultSelectedItems = {[]}
    />{paymentcollectionErrorMessage}

<p className={styles.formlabel}>B/L Clause<span className={styles.required}> *</span></p>

<Dropdown
    selectedKey={blclause}
    placeholder='Select B/L Clause'
    options={[
      { key: 'MOL B/L', text: 'MOL B/L' },
      { key: 'Others', text: 'Others' },
    
      // Add more options as needed
    ]}
    onChange={(event, option) => this.handleDropdownChange(event, option,'blclause')}
  />{blclauseErrorMessage}
{isblothervisible && <div>
<p className={styles.formlabel}>B/L Clause (Other , please specify)<span className={styles.required}> *</span></p>

<TextField value={this.state.blothers} onChange={(event) => this.handleInputChange(event, 'blothers')} />
</div>} {blotherErrorMessage}  

 <p className={styles.formlabel}>Documents (e.g. tender documents / draft contract)</p>

 <div className="mt-5 text-center">
        <label htmlFor="vattachment" className="btn btn-primary text-light" role="button" aria-disabled="false">
          + Add Supporting Documents
        </label>
        <input
          type="file"
          name="file[]"
          accept=""
          id="vattachment"
          style={{ visibility: 'hidden', position: 'absolute' }}
          multiple
          onChange={this.vhandleFileUpload}
        />

      <p id="vfiles-area">
          <span id="vfilesList">
            <span ref={this.filesNamesRef} id="vfiles-names"></span>
          </span>
        </p>
      </div>


 <p className={styles.formlabel}>Applicable Law</p>
      

 <TextField value={this.state.applaw} onChange={(event) => this.handleInputChange(event, 'applaw')} multiline rows={3} />

      </div>
      <div>
      <br /><p className={styles.heading}>Additional Information</p>
      {/* <RichText label="Background" value={this.state.background}  onChange={(text)=>this.onBackgroundTextChange(text)}/> */}
      <p className={styles.formlabel}>Background<span className={styles.required}> *</span></p>
  
      
      <TextField value={this.state.background} onChange={(event) => this.handleInputChange(event, 'background')} multiline rows={3} />{backgroundErrorMessage}
       
      <div id = "background" className="mt-5 text-center">
        <label htmlFor="bgattachment" className="btn btn-primary text-light" role="button" aria-disabled="false">
          + Add Supporting Documents
        </label>
        <input
          type="file"
          name="file[]"
          accept=""
          id="bgattachment"
          style={{ visibility: 'hidden', position: 'absolute' }}
          multiple
          onChange={this.bghandleFileUpload}
        />

        <p id="bgfiles-area">
          <span id="bgfilesList">
            <span ref={this.filesNamesRef} id="bgfiles-names"></span>
          </span>
        </p>
      </div>
      
      {/* <RichText label="Voyage P/L Contribution" value={this.state.voyage}  onChange={(text)=>this.onvoyageTextChange(text)}/>  */}
    
    {/* <RichText label="Others" value={this.state.addothers}  onChange={(text)=>this.onaddothersTextChange(text)}/>  */}
    <p className={styles.formlabel}>Others<span className={styles.required}> *</span></p>

    
    <TextField value={this.state.othercon} onChange={(event) => this.handleInputChange(event, 'othercon')} multiline rows={3} />{otherconErrorMessage}
    <div className="mt-5 text-center">
        <label htmlFor="othersattachment" className="btn btn-primary text-light" role="button" aria-disabled="false">
          + Add Supporting Documents
        </label>
        <input
          type="file"
          name="file[]"
          accept=""
          id="othersattachment"
          style={{ visibility: 'hidden', position: 'absolute' }}
          multiple
          onChange={this.ohandleFileUpload}
        />

<p id="ofiles-area">
          <span id="ofilesList">
            <span ref={this.filesNamesRef} id="ofiles-names"></span>
          </span>
        </p>
      </div>
   
 {console.log("j",this.state.compstr)}
      <p className={styles.formlabel}>Compliance<span className={styles.required}> *</span></p>  
      <div style={{ display: 'inline-flex' }}>
      <Checkbox onChange={this._onCheckboxChange} />
      <span dangerouslySetInnerHTML={{ __html: this.state.compstr }}></span></div>{complianceErrorMessage}

      
    <PeoplePicker
        context={this.props.context as any}
        titleText="Interested Parties (MOLEA)"
        personSelectionLimit={10}
        groupName={""} 
        showtooltip={false}
        ensureUser={true}
        required={false}
        disabled={false}
        onChange={this._getPartiesPeoplePickerItems}
        showHiddenInUI={false}
        principalTypes={[PrincipalType.User]}
        resolveDelay={1000} />                  
    <br/>
    <Stack horizontal verticalAlign="end" className={styles.extpartiesstackContainer }>
          <TextField
            label="Interested Parties (External)"
            value={this.state.newParty}
            styles={textFieldStyles as IStyleFunctionOrObject<ITextFieldStyleProps, ITextFieldStyles>}
            onChange={this._newparty}
           /*  onGetErrorMessage={(value) => {
              if (value && !/^\w+([.-]?\w+)*@\w+([.-]?\w+)*(\.\w{2,3})+$/.test(value)) {
                return 'Please enter a valid email address';
              }
              return '';
            }} */
          />
          <PrimaryButton text="+" onClick={this.handleAddParty} />
          {/* <IconButton onClick= { this.handleAddParty } iconProps={ { iconName: 'Add' } } title='Add Party' /> */}
        </Stack>
    
        <div>
          {interestedPartiesexternal.map((party: any, index: React.Key) => (
            <span key={index}>{party}{index !== interestedPartiesexternal.length - 1 && '; '}</span>
          ))}
        </div>    
        <br/>
        {EmailFieldErrorMessage}
    <Stack horizontal horizontalAlign='end' className={styles.stackContainer}>     
    <PrimaryButton text={buttontext} onClick={() => this._createItem(this.props)} disabled= {isbuttondisbled}/>
    <PrimaryButton text="Cancel"  onClick ={this.cancel}/>
   
    </Stack> 
    <br />
    <div>      
      {isLoading && <Spinner label="Saving, please wait..." size={SpinnerSize.large} />}
      </div>
    {/* {imageFieldErrorMessage} */}
    <br />
    {FormFieldErrorMessage}
    {successMessage}
    
        </div>
      </section>
    );
  }
}


