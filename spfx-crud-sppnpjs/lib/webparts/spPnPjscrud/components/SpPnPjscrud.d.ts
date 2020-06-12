import * as React from 'react';
import { ISpPnPjscrudProps } from './ISpPnPjscrudProps';
import { IEmployees } from './IEmployees';
import { IEmployee } from './IEmployee';
import 'bootstrap/dist/css/bootstrap.min.css';
export default class SpPnPjscrud extends React.Component<ISpPnPjscrudProps, IEmployees> {
    constructor(props: ISpPnPjscrudProps, state: IEmployees);
    render(): React.ReactElement<ISpPnPjscrudProps>;
    private _getFormattedDate;
    private _getState;
    private _getPeoplePickerItems;
    private _onFormatDate;
    private EmployeeOnChange;
    NewEmployee(): void;
    CreateEmployee(): void;
    EditEmployee(item: IEmployee): void;
    componentDidMount(): void;
    private ReadEmployees;
    DeleteEmployee: (Id: number) => void;
}
//# sourceMappingURL=SpPnPjscrud.d.ts.map