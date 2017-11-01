import { BaseEmployee } from './BaseEmployee';

export class Employee {
    public Id?: number;
    public UserName?: string;
    public FullName?: string;
    public UserPosition?: string;
    public Department?: string;
    public Direction?: string;
    public Days?: number;
    public DaysLeftNextYear?: number;
    public Replacements?: Employee[];
}