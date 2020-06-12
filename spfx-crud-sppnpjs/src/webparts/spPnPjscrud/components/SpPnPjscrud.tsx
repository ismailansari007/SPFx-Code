import * as React from 'react';
import styles from './SpPnPjscrud.module.scss';
import { ISpPnPjscrudProps } from './ISpPnPjscrudProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { sp, ItemAddResult, ItemUpdateResult } from 'sp-pnp-js'
import { IEmployees } from './IEmployees';
import { IEmployee } from './IEmployee';
import Select from 'react-select';
// import { sp } from "@pnp/sp";
// import "@pnp/sp/webs";
// import "@pnp/sp/site-users/web";
// import "@pnp/sp/lists";
// import "@pnp/sp/items";
import { IItemAddResult, Items } from "@pnp/sp/items";
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { DateTimePicker, DateConvention, TimeConvention, TimeDisplayControlType } from '@pnp/spfx-controls-react/lib/dateTimePicker';
import { Button, Container, Row, Col } from 'reactstrap';
import 'bootstrap/dist/css/bootstrap.min.css';

export default class SpPnPjscrud extends React.Component<ISpPnPjscrudProps, IEmployees> {
  public constructor(props: ISpPnPjscrudProps, state: IEmployees) {
    super(props);
    this.state = {
      Id: 0,
      EmployeeName: '',
      EmployeeAddress: '',
      State: [],
      selectedState: null,
      button: 'Create Employee',
      Date: null,
      Users: [],
      UsersTitle: [],
      items: []
    };
    this._getPeoplePickerItems = this._getPeoplePickerItems.bind(this);
  }

  public render(): React.ReactElement<ISpPnPjscrudProps> {
    return (
      <div className={styles.spPnPjscrud}>
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.column}>
              <a href="#" onClick={() => this.NewEmployee()} className={styles.button}>
                <span className={styles.label}>New Employee</span>
              </a>
            </div>
          </div>
          <div className={styles.row}>
            <div className={styles.column}>
              <p>
                <label><b>Employee Name :</b> </label>
                <div>
                  <input type="text" name="EmployeeName" onChange={this.EmployeeOnChange.bind(this)} value={this.state.EmployeeName} className={styles["input-control"]} />
                </div>
              </p>
              <p>
                <label><b>Employee Address :</b> </label>
                <div>
                  <textarea name="EmployeeAddress" onChange={this.EmployeeOnChange.bind(this)} value={this.state.EmployeeAddress} className={styles["input-control"]}>
                  </textarea>
                </div>
              </p>
              <p>
                <label><b>State :</b></label>
                <div>
                  <Select
                    className={styles["input-control"]}
                    classNamePrefix="Select State"
                    defaultInputValue={null}
                    isSearchable={true}
                    value={this.state.selectedState}
                    options={this.state.State}
                    onChange={(e: string)=>this.setState({selectedState: e})}
                  />
                </div>
              </p>
              <p>
                <label><b>Users :</b> </label>
                <PeoplePicker
                  context={this.props.context}
                  // titleText="User"
                  personSelectionLimit={3}
                  groupName={""} // Leave this blank in case you want to filter from all users
                  showtooltip={true}
                  isRequired={true}
                  peoplePickerCntrlclassName={styles["input-control"]}
                  disabled={false}
                  ensureUser={true}
                  selectedItems={this._getPeoplePickerItems}
                  showHiddenInUI={false}
                  principalTypes={[PrincipalType.User]}
                  defaultSelectedUsers={this.state.UsersTitle}
                  resolveDelay={1000} />
              </p>
              <p>
                <label><b>Date :</b> </label>
                <div className={styles["input-control"]}>
                  <DateTimePicker label=""
                    showLabels={false}
                    formatDate={this._onFormatDate}
                    dateConvention={DateConvention.Date}
                    isMonthPickerVisible={false}
                    showMonthPickerAsOverlay={true}
                    value={this.state.Date}
                    onChange={(e)=> this.setState({Date: e})}
                    />
                </div>
              </p>
              <a href="#" onClick={() => this.CreateEmployee()} className={styles.button}>
                <span className={styles.label}>{this.state.button}</span>
              </a>
            </div>
          </div>
          <div className={styles.row}>
            <div className={styles.column}>
              <table className={styles.table}>
                <thead>
                  <tr>
                    <th>Employee Name</th>
                    <th>Employee Address</th>
                    <th>Date</th>
                    <th>Users</th>
                    <th>Action</th>
                  </tr>
                </thead>
                <tbody>
                  {
                    this.state.items.map(function (item: IEmployee, key) {
                      return (
                        <tr key={item.Id}>
                          <td>{item.EmployeeName}</td>
                          <td>{item.EmployeeAddress}</td>
                          <td>{this._getFormattedDate(item.Date)}</td>
                          <td>{item.Users ? item.Users.map(user => user.Title).join(', ') : ""}</td>
                          <td>
                            <a href="#" onClick={() => this.EditEmployee(item)} className={styles.button}>
                              <span className={`${styles.label}`}>Edit</span>
                            </a>
                            <a href="#" onClick={() => this.DeleteEmployee(item.Id)} className={`${styles.button}`}>
                              <span className={styles.label}>Delete</span>
                            </a>
                          </td>
                        </tr>
                      )
                    }.bind(this))
                  }
                </tbody>
              </table>
            </div>
          </div>
        </div>
      </div>
    );
  }

  private _getFormattedDate(date: string){
    let formattedDate = "";
    if(date){
      let nDate = new Date(date.split('T')[0]);
      let day = nDate.getDate() < 10 ? "0" + nDate.getDate() : nDate.getDate();
      let month = nDate.getMonth() < 9 ? "0" + (nDate.getMonth() + 1) : (nDate.getDate() + 1);
      formattedDate = day + "/" + month + "/" + nDate.getFullYear();
    }
    return formattedDate;
  }

  private _getState(){
    try{
      sp.web.lists.getByTitle('State').items.select("Id", "Title").get().then(
        (response) => {
          let options = response.map(function(item){
            let option = { "value": item.Id, "label": item.Title };
            return option;
          });
          this.setState({State: options});
        }
      )
    }
    catch(ex){

    }
  }

  private _getPeoplePickerItems(items: any[]) {
    try {
      console.log('Items:', items);
      let userIds = items.map(item => item.id);
      let userLogins = items.map(item => item.secondaryText);
      this.setState({ Users: userIds, UsersTitle: userLogins });
    }
    catch (error) {

    }
  }

  private _onFormatDate = (date: Date): string => {
    let day = date.getDate() < 10 ? "0" + date.getDate() : date.getDate();
    let month = date.getMonth() < 9 ? "0" + (date.getMonth() + 1) : (date.getMonth() + 1);
    return day + '/' + month + '/' + date.getFullYear();
  };

  private EmployeeOnChange(e) {
    let change = {};
    change[e.target.name] = e.target.value;
    console.log(change);
    this.setState(change)
    console.log(this.state);
  }

  public NewEmployee(): void {
    this.setState({ Id: 0, EmployeeName: '', EmployeeAddress: '', button: 'Create Employee', Users: [], UsersTitle: [], Date: null, selectedState: null })
  }

  public CreateEmployee(): void {
    console.log(this.state.selectedState);
    let date = (this.state.Date.getMonth() + 1) + "/" + (this.state.Date.getDate()) + "/" + this.state.Date.getFullYear();
    if (this.state.Id == 0) {
      console.log(this.state.EmployeeName + ' ' + this.state.EmployeeAddress)
      sp.web.lists.getByTitle("Employee").items.add({
        'EmployeeName': this.state.EmployeeName,
        'EmployeeAddress': this.state.EmployeeAddress,
        'State': this.state.selectedState["Id"],
        'UsersId': { 'results': this.state.Users },
        'Date': date
      }).then((result: ItemAddResult): void => {
        this.NewEmployee();
        this.ReadEmployees();
        alert('Employee entry created successfully');
      }, (error: any): void => {
        console.log(error);
      });
    }
    else {
      sp.web.lists.getByTitle("Employee").items.getById(this.state.Id).update({
        'EmployeeName': this.state.EmployeeName,
        'EmployeeAddress': this.state.EmployeeAddress,
        'UsersId': { 'results': this.state.Users },
        'Date': date
      }).then((result: ItemUpdateResult): void => {
        this.NewEmployee();
        this.ReadEmployees();
        alert('Employee entry updated successfully');
      }, (error: any): void => {
        console.log(error);
      });
    }
  }

  public EditEmployee(item: IEmployee): void {
    let date = item.Date ? new Date(item.Date.split("T")[0]) : null;
    let usersIds = item.Users ? item.Users.map(user => user.Id) : [];
    let usersEmails = item.Users ? item.Users.map(user => user.EMail) : [];
    this.setState({ Id: item.Id, EmployeeName: item.EmployeeName, EmployeeAddress: item.EmployeeAddress, Users: usersIds, UsersTitle: usersEmails, Date: date });
    this.setState({ 'button': "Update Employee" });
  }

  public componentDidMount() {
    this._getState();
    this.ReadEmployees();
  }

  private ReadEmployees(): void {
    sp.web.lists.getByTitle('Employee').items.select("Id", "EmployeeName", "EmployeeAddress", "Date", "Users/Id", "Users/Title", "Users/EMail").expand("Users").get().then(
      (response) => {
        console.log(response);
        // let employeeCollection = response.map(item => new IEmployees(item));
        this.setState({ items: response });
      }
    )
  }

  public DeleteEmployee = (Id: number): void => {
    console.log(Id);
    if (confirm("Are you sure you want to delete this item")) {
      sp.web.lists.getByTitle("Employee").items.getById(Id).delete().then((result: any): void => {
        this.ReadEmployees();
        alert('Employee entry deleted successfully');
      }, (error: any): void => {
        console.log(error);
      });
    }
  }
}
