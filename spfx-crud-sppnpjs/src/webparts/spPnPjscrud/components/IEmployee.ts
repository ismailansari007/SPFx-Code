import { Users } from "./Users";

export interface IEmployee{
    Id: number;
    EmployeeName: string;
    EmployeeAddress: string;
    Users: Users[];
    Date: string;
}