import * as React from 'react';
import { IStaffCrudProps } from './IStaffCrudProps';
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import styles from './StaffCrud.module.scss';

interface IListItem {
  ID: number;
  Title: string;
  DOB: any;
  Address:String;
  Department:any;
  ContactNo:number;
  IsMarried:boolean;
  Salary:number;
  Manager:{
    Title: string;
    EMail: string;
  };

}
interface IListItems {
  AllItems: IListItem[];

  
  listTitle: string;
  listAddress:any;
  listDOB: number;
  listDepartment:any;
  listSelectedID: number;
  listContactNo:any;
  listSalary:any;
  listIsMarried:any;
}

export default class StaffCrud extends React.Component<IStaffCrudProps, IListItems> {

  
    constructor(props: IStaffCrudProps, state: IListItems) {
      super(props);
      this.state = {
        AllItems: [],
        listTitle: undefined,
        listDOB: 0,
        listSelectedID: 0,
        listAddress:undefined,
        listDepartment:undefined,
        listContactNo:"",
        listSalary:"",
        listIsMarried:undefined,
      };
    }
    componentDidMount() {
      this.getListItems();
    }
     // Get items
  public getListItems = () => {
    // let selectColumns =  `Manager/Title`;
    // let expandColumns =  `Manager`;$expand=${expandColumns}

    let requestURL = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.props.listName}')/items`;
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
    // let expandColumns =  `Manager`;&$expand=${expandColumns}
    // let selectColumns =  `Manager/Title`;
    

    let requestURL = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.props.listName}')/items(${itemID})`;

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
  
  // let expandColumns =  `Manager`;$expand=${expandColumns}
  // let selectColumns =  `Manager/Title`;

  let requestURL = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.props.listName}')/items`;

  const body: string = JSON.stringify({
    Title: this.state.listTitle,
    DOB: this.state.listDOB,
    Address:this.state.listAddress,
    Department:this.state.listDepartment,
    ContactNo:this.state.listContactNo,
    Salary:this.state.listSalary,
    IsMarried:this.state.listIsMarried,
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
  
  // let expandColumns =  `Manager`;&$expand=${expandColumns}
  // let selectColumns =  `Manager/Title`;

  let requestURL = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.props.listName}')/items(${itemID})`;

  const body: string = JSON.stringify({
    Title: this.state.listTitle,
    DOB: this.state.listDOB,
    Address:this.state.listAddress,
    Department:this.state.listDepartment,
    ContactNo:this.state.listContactNo,
    Salary:this.state.listSalary,
    IsMarried:this.state.listIsMarried,
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

  public render(): React.ReactElement<IStaffCrudProps> {
    

    return (
      <div className={styles["spfx-crud"]}>
        <input 
          value={this.state.listTitle}
          type="text"
          name=""
          id="lsTitle"
          placeholder="Title"
          onChange={(e) => {
            this.setState({
              listTitle: e.currentTarget.value,
            });
            // console.log(this.state.listTitle);
          }}
        />
        <input 
          value={this.state.listAddress}
          type="text"
          name=""
          id="lsAddress"
          placeholder="Address"
          onChange={(e) => {
            this.setState({
              listAddress: e.currentTarget.value,
            });
            
          }}
        />

         <input 
          value={this.state.listDOB}
          type="date"
          name=""
          id="lsDOB"
          placeholder="DOB"
          onChange={(e) => {
            this.setState({
              listDOB: e.currentTarget.value as any,
            });
          }}
        />
       <select 
          value={this.state.listDepartment}
          placeholder='Department'
          id="IsDepartment" name="Department"
             onChange={(e) => {
            this.setState({
              listDepartment: e.currentTarget.value as any,
            });
            
          }}>

          <option value="Sales">Sales</option>
          <option value="Marketing">Marketing</option>
          <option value="IT">IT</option>
        </select>
        <input 
          value={this.state.listContactNo}
          type="text"
          name=""
          id="lsNumber"
          placeholder="ContactNo"
          onChange={(e) => {
            this.setState({
              listContactNo: e.currentTarget.value as any,
            });
          }}
        />
        < input
          value={this.state.listIsMarried}
          type="checkbox"
          name=""
          id="lsMarried"
          placeholder="checkbox"
          onChange={(e) => {
            this.setState({
              listIsMarried: e.currentTarget.value as any,
            });
            
          }}
          // < label for="Is Married"> Yes</label>
          // <label for="Is Married"> No</label>
          />
        

        <input 
          value={this.state.listSalary}
          type="text"
          name=""
          id="lsNumber"
          placeholder="Salary"
          onChange={(e) => {
            this.setState({
              listSalary: e.currentTarget.value as any,
            });
          }}
        />
        
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

      <table>
      <th>Title</th>
      <th>DOB</th>
      <th>Address</th>
      <th>Department</th>
      <th>Contact No</th>
      <th>Salary</th>
      <th>Is Married</th>
      {this.state.AllItems.map((emp) => {
        return (
          <tr>
            <td>{emp.Title}</td>
            <td>{emp.DOB}</td>
            <td>{emp.Address}</td>
            <td>{emp.Department}</td>
            <td>{emp.ContactNo}</td>
            <td>{emp.Salary}</td>
            <td><td>{" "}{emp.IsMarried?"Yes":"No"}</td></td>
            <td>
              <button
                onClick={() => {
                  this.setState({
                    listTitle: emp.Title,
                    listDOB: emp.DOB,
                    listSelectedID: emp.ID,
                    listAddress:emp.Address,
                    listDepartment:emp.Department,
                    listContactNo:emp.ContactNo,
                    listSalary:emp.Salary,
                    listIsMarried:emp.IsMarried,
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
    );
  }
}
