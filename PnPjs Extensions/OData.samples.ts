import { SPOData } from "./OData";

//Sample lists from SP

export interface IDepartment {
    Id: number;
    Manager: IDepartment;
    Alias: string;
    HasSharedMailbox: boolean;
    DepartmentNumber: number;
}

export interface IEmployee {
    Id: number;
    Age: number;
    Firstname: string;
    Lastname: string;
    Employed: boolean;
    Department: IDepartment;
    DepartmentId: number
    Manager: IEmployee;
    Created: Date;
    Emails: string[];
    SecondaryDepartment: IDepartment[];
    SecondaryDepartmentId: number[];
}

//Sample queries

// Firstname eq 'David'
const Query1 = SPOData.Where<IEmployee>().TextField("Firstname").EqualTo("David").ToString();

// (DepartmentId eq 1 and Employed eq 1 and Age le 30 and Department/Alias eq 'Consulting' and Department/DepartmentNumber gt 50 and startswith(Firstname, 'D') and substringof('oft', Lastname) and startswith(Manager/Firstname, 'B'))
const Query2 = SPOData.Where<IEmployee>().All([
    SPOData.Where<IEmployee>().LookupIdField("DepartmentId").EqualTo(1),
    SPOData.Where<IEmployee>().BooleanField("Employed").IsTrue(),
    SPOData.Where<IEmployee>().NumberField("Age").LessThanOrEqualTo(30),
    SPOData.Where<IEmployee>().LookupField("Department").TextField("Alias").EqualTo("Consulting"),
    SPOData.Where<IEmployee>().LookupField("Department").NumberField("DepartmentNumber").GreaterThan(50),
    SPOData.Where<IEmployee>().TextField("Firstname").StartsWith("D"),
    SPOData.Where<IEmployee>().TextField("Lastname").Contains("oft"),
    SPOData.Where<IEmployee>().LookupField("Manager").TextField("Firstname").StartsWith("B"),
]).ToString();

// DepartmentId eq 1 and Employed eq 1 and Age le 30 and Department/Alias eq 'Consulting' and Department/DepartmentNumber gt 50 and startswith(Firstname, 'D') and substringof('oft', Lastname) and startswith(Manager/Firstname, 'B')
const Query3 = SPOData.Where<IEmployee>()
    .LookupIdField("DepartmentId").EqualTo(1)
    .And()
    .BooleanField("Employed").IsTrue()
    .And()
    .NumberField("Age").LessThanOrEqualTo(30)
    .And()
    .LookupField("Department").TextField("Alias").EqualTo("Consulting")
    .And()
    .LookupField("Department").NumberField("DepartmentNumber").GreaterThan(50)
    .And()
    .TextField("Firstname").StartsWith("D")
    .And()
    .TextField("Lastname").Contains("oft")
    .And()
    .LookupField("Manager").TextField("Firstname").StartsWith("B")
    .ToString();

// (Department/DepartmentNumber eq 50 or Department/DepartmentNumber eq 51 or Department/DepartmentNumber eq 52 or Department/DepartmentNumber eq 53)
const Query4 = SPOData.Where<IEmployee>().Some([
    SPOData.Where<IEmployee>().LookupField("Department").NumberField("DepartmentNumber").EqualTo(50),
    SPOData.Where<IEmployee>().LookupField("Department").NumberField("DepartmentNumber").EqualTo(51),
    SPOData.Where<IEmployee>().LookupField("Department").NumberField("DepartmentNumber").EqualTo(52),
    SPOData.Where<IEmployee>().LookupField("Department").NumberField("DepartmentNumber").EqualTo(53),
]).ToString();

// Department/DepartmentNumber eq 50 or Department/DepartmentNumber eq 51 or Department/DepartmentNumber eq 52 or Department/DepartmentNumber eq 53
const Query5 = SPOData.Where<IEmployee>()
    .LookupField("Department").NumberField("DepartmentNumber").EqualTo(50)
    .Or()
    .LookupField("Department").NumberField("DepartmentNumber").EqualTo(51)
    .Or()
    .LookupField("Department").NumberField("DepartmentNumber").EqualTo(52)
    .Or()
    .LookupField("Department").NumberField("DepartmentNumber").EqualTo(53)
    .ToString();

// (Department/DepartmentNumber eq 50 or Department/DepartmentNumber eq 51 or Department/DepartmentNumber eq 52 or Department/DepartmentNumber eq 53)
const Query6 = SPOData.Where<IEmployee>().LookupField("Department").NumberField("DepartmentNumber").In([50, 51, 52, 53]).ToString();

// (Created gt '2023-11-22T23:00:00.000Z' and Created lt '2023-11-23T22:59:59.999Z')
const Query7 = SPOData.Where<IEmployee>().DateField("Created").IsToday().ToString();

// (Created gt '2022-12-31T23:00:00.000Z' and Created lt '2023-12-31T22:59:59.999Z')
const Query8 = SPOData.Where<IEmployee>().DateField("Created").IsBetween(new Date(2023, 0, 1, 0, 0, 0, 0), new Date(2023, 11, 31, 23, 59, 59, 999)).ToString();



SPOData.Where<IEmployee>().All([
    SPOData.Where().DateField("Created").IsBetween(new Date(2023, 0, 1, 0, 0, 0, 0), new Date(2023, 11, 31, 23, 59, 59, 999)),
    SPOData.Where().LookupField("Department").TextField("Alias").Contains("Consulting"),
]).ToString();

console.log(Query1);
console.log("\n");
console.log(Query2);
console.log("\n");
console.log(Query3);
console.log("\n");
console.log(Query4);
console.log("\n");
console.log(Query5);
console.log("\n");
console.log(Query6);
console.log("\n");
console.log(Query7);
console.log("\n");
console.log(Query8);
