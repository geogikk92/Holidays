import { BaseEmployee } from './BaseEmployee';
import { Employee } from './Employee';

export class Holiday {
    public Id: number;
    public Title: string;
    public Address: string;
    public Days: number;
    public DateFrom: Date;
    public DateTo: Date;
    public Replacement: BaseEmployee;
    public Description: string;
    public TypeRequest: string;
    public Type: string;
    public SubType: string;
    public Mobile: string;
    public TotalAllowedDays: number;
    public Employee: Employee;
    public Requestor: BaseEmployee;
    public SiteUrl?: string;
    public Status?: string;
}
