var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
import * as React from 'react';
import styles from './SpPnPjscrud.module.scss';
import { sp } from 'sp-pnp-js';
import Select from 'react-select';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { DateTimePicker, DateConvention } from '@pnp/spfx-controls-react/lib/dateTimePicker';
import 'bootstrap/dist/css/bootstrap.min.css';
var SpPnPjscrud = /** @class */ (function (_super) {
    __extends(SpPnPjscrud, _super);
    function SpPnPjscrud(props, state) {
        var _this = _super.call(this, props) || this;
        _this._onFormatDate = function (date) {
            var day = date.getDate() < 10 ? "0" + date.getDate() : date.getDate();
            var month = date.getMonth() < 9 ? "0" + (date.getMonth() + 1) : (date.getMonth() + 1);
            return day + '/' + month + '/' + date.getFullYear();
        };
        _this.DeleteEmployee = function (Id) {
            console.log(Id);
            if (confirm("Are you sure you want to delete this item")) {
                sp.web.lists.getByTitle("Employee").items.getById(Id).delete().then(function (result) {
                    _this.ReadEmployees();
                    alert('Employee entry deleted successfully');
                }, function (error) {
                    console.log(error);
                });
            }
        };
        _this.state = {
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
        _this._getPeoplePickerItems = _this._getPeoplePickerItems.bind(_this);
        return _this;
    }
    SpPnPjscrud.prototype.render = function () {
        var _this = this;
        return (React.createElement("div", { className: styles.spPnPjscrud },
            React.createElement("div", { className: styles.container },
                React.createElement("div", { className: styles.row },
                    React.createElement("div", { className: styles.column },
                        React.createElement("a", { href: "#", onClick: function () { return _this.NewEmployee(); }, className: styles.button },
                            React.createElement("span", { className: styles.label }, "New Employee")))),
                React.createElement("div", { className: styles.row },
                    React.createElement("div", { className: styles.column },
                        React.createElement("p", null,
                            React.createElement("label", null,
                                React.createElement("b", null, "Employee Name :"),
                                " "),
                            React.createElement("div", null,
                                React.createElement("input", { type: "text", name: "EmployeeName", onChange: this.EmployeeOnChange.bind(this), value: this.state.EmployeeName, className: styles["input-control"] }))),
                        React.createElement("p", null,
                            React.createElement("label", null,
                                React.createElement("b", null, "Employee Address :"),
                                " "),
                            React.createElement("div", null,
                                React.createElement("textarea", { name: "EmployeeAddress", onChange: this.EmployeeOnChange.bind(this), value: this.state.EmployeeAddress, className: styles["input-control"] }))),
                        React.createElement("p", null,
                            React.createElement("label", null,
                                React.createElement("b", null, "State :")),
                            React.createElement("div", null,
                                React.createElement(Select, { className: styles["input-control"], classNamePrefix: "Select State", defaultInputValue: null, isSearchable: true, value: this.state.selectedState, options: this.state.State, onChange: function (e) { return _this.setState({ selectedState: e }); } }))),
                        React.createElement("p", null,
                            React.createElement("label", null,
                                React.createElement("b", null, "Users :"),
                                " "),
                            React.createElement(PeoplePicker, { context: this.props.context, 
                                // titleText="User"
                                personSelectionLimit: 3, groupName: "", showtooltip: true, isRequired: true, peoplePickerCntrlclassName: styles["input-control"], disabled: false, ensureUser: true, selectedItems: this._getPeoplePickerItems, showHiddenInUI: false, principalTypes: [PrincipalType.User], defaultSelectedUsers: this.state.UsersTitle, resolveDelay: 1000 })),
                        React.createElement("p", null,
                            React.createElement("label", null,
                                React.createElement("b", null, "Date :"),
                                " "),
                            React.createElement("div", { className: styles["input-control"] },
                                React.createElement(DateTimePicker, { label: "", showLabels: false, formatDate: this._onFormatDate, dateConvention: DateConvention.Date, isMonthPickerVisible: false, showMonthPickerAsOverlay: true, value: this.state.Date, onChange: function (e) { return _this.setState({ Date: e }); } }))),
                        React.createElement("a", { href: "#", onClick: function () { return _this.CreateEmployee(); }, className: styles.button },
                            React.createElement("span", { className: styles.label }, this.state.button)))),
                React.createElement("div", { className: styles.row },
                    React.createElement("div", { className: styles.column },
                        React.createElement("table", { className: styles.table },
                            React.createElement("thead", null,
                                React.createElement("tr", null,
                                    React.createElement("th", null, "Employee Name"),
                                    React.createElement("th", null, "Employee Address"),
                                    React.createElement("th", null, "Date"),
                                    React.createElement("th", null, "Users"),
                                    React.createElement("th", null, "Action"))),
                            React.createElement("tbody", null, this.state.items.map(function (item, key) {
                                var _this = this;
                                return (React.createElement("tr", { key: item.Id },
                                    React.createElement("td", null, item.EmployeeName),
                                    React.createElement("td", null, item.EmployeeAddress),
                                    React.createElement("td", null, this._getFormattedDate(item.Date)),
                                    React.createElement("td", null, item.Users ? item.Users.map(function (user) { return user.Title; }).join(', ') : ""),
                                    React.createElement("td", null,
                                        React.createElement("a", { href: "#", onClick: function () { return _this.EditEmployee(item); }, className: styles.button },
                                            React.createElement("span", { className: "" + styles.label }, "Edit")),
                                        React.createElement("a", { href: "#", onClick: function () { return _this.DeleteEmployee(item.Id); }, className: "" + styles.button },
                                            React.createElement("span", { className: styles.label }, "Delete")))));
                            }.bind(this)))))))));
    };
    SpPnPjscrud.prototype._getFormattedDate = function (date) {
        var formattedDate = "";
        if (date) {
            var nDate = new Date(date.split('T')[0]);
            var day = nDate.getDate() < 10 ? "0" + nDate.getDate() : nDate.getDate();
            var month = nDate.getMonth() < 9 ? "0" + (nDate.getMonth() + 1) : (nDate.getDate() + 1);
            formattedDate = day + "/" + month + "/" + nDate.getFullYear();
        }
        return formattedDate;
    };
    SpPnPjscrud.prototype._getState = function () {
        var _this = this;
        try {
            sp.web.lists.getByTitle('State').items.select("Id", "Title").get().then(function (response) {
                var options = response.map(function (item) {
                    var option = { "value": item.Id, "label": item.Title };
                    return option;
                });
                _this.setState({ State: options });
            });
        }
        catch (ex) {
        }
    };
    SpPnPjscrud.prototype._getPeoplePickerItems = function (items) {
        try {
            console.log('Items:', items);
            var userIds = items.map(function (item) { return item.id; });
            var userLogins = items.map(function (item) { return item.secondaryText; });
            this.setState({ Users: userIds, UsersTitle: userLogins });
        }
        catch (error) {
        }
    };
    SpPnPjscrud.prototype.EmployeeOnChange = function (e) {
        var change = {};
        change[e.target.name] = e.target.value;
        console.log(change);
        this.setState(change);
        console.log(this.state);
    };
    SpPnPjscrud.prototype.NewEmployee = function () {
        this.setState({ Id: 0, EmployeeName: '', EmployeeAddress: '', button: 'Create Employee', Users: [], UsersTitle: [], Date: null, selectedState: null });
    };
    SpPnPjscrud.prototype.CreateEmployee = function () {
        var _this = this;
        console.log(this.state.selectedState);
        var date = (this.state.Date.getMonth() + 1) + "/" + (this.state.Date.getDate()) + "/" + this.state.Date.getFullYear();
        if (this.state.Id == 0) {
            console.log(this.state.EmployeeName + ' ' + this.state.EmployeeAddress);
            sp.web.lists.getByTitle("Employee").items.add({
                'EmployeeName': this.state.EmployeeName,
                'EmployeeAddress': this.state.EmployeeAddress,
                'State': this.state.selectedState["Id"],
                'UsersId': { 'results': this.state.Users },
                'Date': date
            }).then(function (result) {
                _this.NewEmployee();
                _this.ReadEmployees();
                alert('Employee entry created successfully');
            }, function (error) {
                console.log(error);
            });
        }
        else {
            sp.web.lists.getByTitle("Employee").items.getById(this.state.Id).update({
                'EmployeeName': this.state.EmployeeName,
                'EmployeeAddress': this.state.EmployeeAddress,
                'UsersId': { 'results': this.state.Users },
                'Date': date
            }).then(function (result) {
                _this.NewEmployee();
                _this.ReadEmployees();
                alert('Employee entry updated successfully');
            }, function (error) {
                console.log(error);
            });
        }
    };
    SpPnPjscrud.prototype.EditEmployee = function (item) {
        var date = item.Date ? new Date(item.Date.split("T")[0]) : null;
        var usersIds = item.Users ? item.Users.map(function (user) { return user.Id; }) : [];
        var usersEmails = item.Users ? item.Users.map(function (user) { return user.EMail; }) : [];
        this.setState({ Id: item.Id, EmployeeName: item.EmployeeName, EmployeeAddress: item.EmployeeAddress, Users: usersIds, UsersTitle: usersEmails, Date: date });
        this.setState({ 'button': "Update Employee" });
    };
    SpPnPjscrud.prototype.componentDidMount = function () {
        this._getState();
        this.ReadEmployees();
    };
    SpPnPjscrud.prototype.ReadEmployees = function () {
        var _this = this;
        sp.web.lists.getByTitle('Employee').items.select("Id", "EmployeeName", "EmployeeAddress", "Date", "Users/Id", "Users/Title", "Users/EMail").expand("Users").get().then(function (response) {
            console.log(response);
            // let employeeCollection = response.map(item => new IEmployees(item));
            _this.setState({ items: response });
        });
    };
    return SpPnPjscrud;
}(React.Component));
export default SpPnPjscrud;
//# sourceMappingURL=SpPnPjscrud.js.map