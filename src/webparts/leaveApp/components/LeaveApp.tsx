import * as React from 'react';
import styles from './LeaveApp.module.scss';
import { ILeaveAppProps } from './ILeaveAppProps';
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import * as moment from "moment";
// import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";

interface IListItem {
  ID: number;
  Title: string;
  StartDate: any;
  EndDate:any;
  RequireManagerApproval:String
  Reason:string;
  Status:any;
  ManagerEmail:{
    Title: string;
    EMail: string;
  };

}
interface IListItems {
  AllItems: IListItem[];

  listTitle: string;
  listStartDate:number;
  listEndDate: any;
  listRequireManagerApproval:any;
  listSelectedID: number;
  listReason:any;
  listStatus:any;
  // listManagerEmail:any
  
}


export default class LeaveApp extends React.Component<ILeaveAppProps, IListItems> {

  constructor(props: ILeaveAppProps, state: IListItems) {
    super(props);
    this.state = {
      AllItems: [],
      listTitle: undefined,
      listStartDate:0,
      listEndDate: 0,
      listRequireManagerApproval:undefined,
      listSelectedID: 0,
      listReason:undefined,
      listStatus:undefined,
      // listManagerEmail:{
      //     Title:undefined,
      //     EMail:undefined
      //   }
    };
  }
  componentDidMount() {
    this.getListItems();
  }
  // Get items
  public getListItems = () => {
   
    let selectColumns ="Title,ID,Status,StartDate,EndDate,RequireManagerApproval,Reason,ManagerEmail/Title,ManagerEmail/EMail";
    let expandColumns = "ManagerEmail";
  // Filter applied because we need to set only approved kudos
  // let filterBy = `Status eq 'Approved'`;
  // let top = this.props.numberOfKudosToShow;

    let requestURL = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.props.listName}')/items?$select=${selectColumns}&$expand=${expandColumns}`
    // let requestURL = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.props.listName}')/items`;
    this.props.context.spHttpClient
      .get(requestURL, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        if (response.ok) {
          return response.json();
        }
      })
      .then((i) => {
        if (i == undefined) {
        } else {
          this.setState({
            AllItems: i.value,
          });
          console.log(this.state.AllItems);
        }
      });
  };
    // Delete item
    public deleteItem = (itemID: number) => {
      alert("this is delete");
      let selectColumns ="Title,ID,Status,StartDate,EndDate,RequireManagerApproval,Reason,ManagerEmail/Title,ManagerEmail/EMail";
      let expandColumns = "ManagerEmail";
      
  
      // let requestURL = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.props.listName}')/items(${itemID})`;
      let requestURL = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.props.listName}')/items(${itemID})?$select=${selectColumns}&$expand=${expandColumns}`
      this.props.context.spHttpClient
        .post(requestURL, SPHttpClient.configurations.v1, {
          headers: {
            Accept: "application/json;odata=nometadata",
            "Content-type": "application/json;odata=verbose",
            "odata-version": "",
            "IF-MATCH": "*",
            "X-HTTP-Method": "DELETE",
          },
        })
        .then((response: SPHttpClientResponse) => {
          if (response.ok) {
            alert(`Item ID: ${itemID} deleted successfully!`);
            this.getListItems();
          } else {
            alert(`Something went wrong!`);
            console.log(response.json());
          }
        });
    };
    // Add item
public addItemInList = () => {
  
  let selectColumns ="Title,ID,Status,StartDate,EndDate,RequireManagerApproval,Reason,ManagerEmail/Title,ManagerEmail/EMail";
  let expandColumns = "ManagerEmail";

  // let requestURL = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.props.listName}')/items`;
  let requestURL = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.props.listName}')/items?$select=${selectColumns}&$expand=${expandColumns}`

  const body: string = JSON.stringify({
    Title: this.state.listTitle,
    ID:this.state.listSelectedID,
    StartDate:this.state.listStartDate,
    EndDate:this.state.listEndDate,
    RequireManagerApproval:this.state.listRequireManagerApproval,
    Reason:this.state.listReason,
    Status:this.state.listStatus,
    // ManagerEmail:this.state.listManagerEmail,

  });

  this.props.context.spHttpClient
    .post(requestURL, SPHttpClient.configurations.v1, {
      headers: {
        Accept: "application/json;odata=nometadata",
        "Content-type": "application/json;odata=nometadata",
        "odata-version": "",
      },
      body: body,
    })
    .then((response: SPHttpClientResponse) => {
      if (response.ok) {
        alert(`Item added successfully!`);
        this.getListItems();
      } else {
        alert(`Something went wrong!`);
        console.log(response.json());
      }
    });
};
// Update item
public updateItemInList = (itemID: number) => {
  let selectColumns ="Title,ID,Status,StartDate,EndDate,RequireManagerApproval,Reason,ManagerEmail/Title,ManagerEmail/EMail";
  let expandColumns = "ManagerEmail";
 

  // let requestURL = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.props.listName}')/items(${itemID})`;
  let requestURL = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.props.listName}')/items(${itemID})?$select=${selectColumns}&$expand=${expandColumns}`

  const body: string = JSON.stringify({
    Title: this.state.listTitle,
    ID:this.state.listSelectedID,
    StartDate:this.state.listStartDate,
    EndDate:this.state.listEndDate,
    RequireManagerApproval:this.state.listRequireManagerApproval,
    Reason:this.state.listReason,
    Status:this.state.listStatus,
    // ManagerEmail:this.state.listManagerEmail,
  });

  this.props.context.spHttpClient
    .post(requestURL, SPHttpClient.configurations.v1, {
      headers: {
        Accept: "application/json;odata=nometadata",
        "Content-type": "application/json;odata=nometadata",
        "odata-version": "",
        "IF-MATCH": "*",
        "X-HTTP-Method": "MERGE",
      },
      body: body,
    })
    .then((response: SPHttpClientResponse) => {
      if (response.ok) {
        alert(`Item updated successfully!`);
        this.getListItems();
      } else {
        alert(`Something went wrong!`);
        console.log(response.json());
      }
    });
};
// private _getPeoplePickerItems(items: any[]) {
//   console.log('Items:', items);
// }
  public render(): React.ReactElement<ILeaveAppProps> {
   
    return (
      
      <div className={styles["leave-crud"]}>
      <div><h3 className ={styles.heading}> {this.props.listName}</h3></div>

      <label htmlFor="Employee Name">Employee Name:</label><br></br>
        <input 
          value={this.state.listTitle}
          type="text"
          name=""
          id="lsTitle"
          placeholder="Employee Name"
          onChange={(e) => {
            this.setState({
              listTitle: e.currentTarget.value,
            });
            // console.log(this.state.listTitle);
          }}
        /><br></br><br></br>

    <label htmlFor="StartDate">Start Date:</label><br></br>
         <input 
          value={this.state.listStartDate}
          type="date"
          name=""
          id="lsStartDate"
          placeholder="Start Date"
          onChange={(e) => {
            this.setState({
              listStartDate: e.currentTarget.value as any,
            });
          }}
        /><br></br><br></br>

<label htmlFor="EndDate">End Date:</label><br></br>
         <input 
          value={this.state.listEndDate}
          type="date"
          name=""
          id="lsEndDate"
          placeholder="End Date"
          onChange={(e) => {
            this.setState({
              listEndDate: e.currentTarget.value as any,
            });
          }}
        /><br></br><br></br> 

 <label htmlFor="RequireManagerApproval">Require Manager Approval:</label><br></br>
        <input 
          value={this.state.listRequireManagerApproval}
          type="text"
          name=""
          id="lsRequire"
          placeholder="Require Manager Approval"
          onChange={(e) => {
            this.setState({
              listRequireManagerApproval: e.currentTarget.value,
            });
            
          }}
        /><br></br><br></br>

<label htmlFor="Reason">Reason:</label><br></br>
        <input 
          value={this.state.listReason}
          type="text"
          name=""
          id="lsReason"
          placeholder="Reason"
          onChange={(e) => {
            this.setState({
              listReason: e.currentTarget.value,
            });
           
          }}
        /><br></br><br></br> 
  {/* <label htmlFor="Status">Status:</label><br></br>
       <select 
          value={this.state.listStatus}
          placeholder='Status'
          id="IsStatus" name="Status"
             onChange={(e) => {
            this.setState({
              listStatus: e.currentTarget.value as any,
            });
            
          }}>

          <option value="Pending">Pending</option>
          <option value="Approved">Approved</option>
          <option value="Rejected">Rejected</option>
          </select><br></br><br></br>
           <label htmlFor="Manager">Manager:</label><br></br> */}
        {/* <input 
          value={this.state.listManagerEmail}
          type="text"
          name=""
          id="lsManager"
          placeholder="Manager"
          onChange={(e) => {
            this.setState({
              listManagerEmail: e.currentTarget.value as any,
            });
            
          }}
        /><br></br><br></br> */}

         {/* <div>
         <label htmlFor="Manager">Manager Email:</label><br></br> 
                  <PeoplePicker
                    context={this.props.context}
                    personSelectionLimit={1}
                    // defaultSelectedUsers={this.state.listManagerEmail===""?[]:this.state.listManagerEmail}
                    required={false}
                    onChange={this._getPeoplePickerItems}
                    defaultSelectedUsers={[this.state.listManagerEmail?this.state.listManagerEmail:""]}
                    showHiddenInUI={false}
                    principalTypes={[PrincipalType.User]}
                    resolveDelay={1000}
                    ensureUser={true}
                  />
                </div> */}

          <button 
          onClick={() => {
            this.addItemInList();
          }}
        >
          Submit
        </button>
       
        <button 
          onClick={() => {
            this.updateItemInList(this.state.listSelectedID);
          }}
        >
          Update
        </button>
        <hr />
        <hr />

        <div className={styles.container}>
      <table>
      <th>Title</th>
      <th>Start Date</th>
      <th>End Date</th>
      <th>R. M. Approval</th>
      <th>Reason</th>
      <th>Status</th>
      <th>Manager Email</th>
      <th></th>
      <th></th>
      
      {this.state.AllItems.map((emp) => {
        return (
          <tr>
            <td>{emp.Title}</td>
            <td>{moment(emp.StartDate).format("LL")}</td>
            <td>{moment(emp.EndDate).format("LL")}</td>
            <td>{emp.RequireManagerApproval}</td>
            <td>{emp.Reason}</td>
            <td>{emp.Status}</td>
            {/* <td>{emp.ManagerEmail.Title}</td> */}
            
            {/* <td>{emp.ManagerEmail==undefined?"":emp.ManagerEmail.EMail}</td> */}
            <td>
              <button
                onClick={() => {
                  this.setState({
                    listTitle: emp.Title,
                    listStartDate: emp.StartDate,
                    listSelectedID: emp.ID,
                    listEndDate:emp.EndDate,
                    listRequireManagerApproval:emp.RequireManagerApproval,
                    listReason:emp.Reason,
                    listStatus:emp.Status,
                    // listManagerEmail:emp.ManagerEmail.EMail
                    
                  });
                }}
              >
                Edit
              </button>
            </td>
            <td>
              <button
                onClick={() => {
                  this.deleteItem(emp.ID);
                }}
              >
                Delete
              </button>
            </td>
          </tr>
        );
      })}
    </table>
  </div>  
      </div>
    );
  }
}
