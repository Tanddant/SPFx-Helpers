import { OData } from "./OData";

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
}

//Sample queries

// Firstname eq 'David'
const Query1 = OData.Where<IEmployee>().TextField("Firstname").EqualTo("David").ToString();

// (DepartmentId eq 1 and Employed eq 1 and Age le 30 and Department/Alias eq 'Consulting' and Department/DepartmentNumber gt 50 and startswith(Firstname, 'D') and substringof('oft', Lastname) and startswith(Manager/Firstname, 'B'))
const Query2 = OData.Where<IEmployee>().All([
    OData.Where<IEmployee>().LookupIdField("DepartmentId").EqualTo(1),
    OData.Where<IEmployee>().BooleanField("Employed").IsTrue(),
    OData.Where<IEmployee>().NumberField("Age").LessThanOrEqualTo(30),
    OData.Where<IEmployee>().LookupField("Department").TextField("Alias").EqualTo("Consulting"),
    OData.Where<IEmployee>().LookupField("Department").NumberField("DepartmentNumber").GreaterThan(50),
    OData.Where<IEmployee>().TextField("Firstname").StartsWith("D"),
    OData.Where<IEmployee>().TextField("Lastname").Contains("oft"),
    OData.Where<IEmployee>().LookupField("Manager").TextField("Firstname").StartsWith("B"),
]).ToString();

// DepartmentId eq 1 and Employed eq 1 and Age le 30 and Department/Alias eq 'Consulting' and Department/DepartmentNumber gt 50 and startswith(Firstname, 'D') and substringof('oft', Lastname) and startswith(Manager/Firstname, 'B')
const Query3 = OData.Where<IEmployee>()
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
const Query4 = OData.Where<IEmployee>().Some([
    OData.Where<IEmployee>().LookupField("Department").NumberField("DepartmentNumber").EqualTo(50),
    OData.Where<IEmployee>().LookupField("Department").NumberField("DepartmentNumber").EqualTo(51),
    OData.Where<IEmployee>().LookupField("Department").NumberField("DepartmentNumber").EqualTo(52),
    OData.Where<IEmployee>().LookupField("Department").NumberField("DepartmentNumber").EqualTo(53),
]).ToString();

// Department/DepartmentNumber eq 50 or Department/DepartmentNumber eq 51 or Department/DepartmentNumber eq 52 or Department/DepartmentNumber eq 53
const Query5 = OData.Where<IEmployee>()
    .LookupField("Department").NumberField("DepartmentNumber").EqualTo(50)
    .Or()
    .LookupField("Department").NumberField("DepartmentNumber").EqualTo(51)
    .Or()
    .LookupField("Department").NumberField("DepartmentNumber").EqualTo(52)
    .Or()
    .LookupField("Department").NumberField("DepartmentNumber").EqualTo(53)
    .ToString();

// (Department/DepartmentNumber eq 50 or Department/DepartmentNumber eq 51 or Department/DepartmentNumber eq 52 or Department/DepartmentNumber eq 53)
const Query6 = OData.Where<IEmployee>().LookupField("Department").NumberField("DepartmentNumber").In([50, 51, 52, 53]).ToString();

// (Created gt '2023-11-22T23:00:00.000Z' and Created lt '2023-11-23T22:59:59.999Z')
const Query7 = OData.Where<IEmployee>().DateField("Created").IsToday().ToString();

// (Created gt '2022-12-31T23:00:00.000Z' and Created lt '2023-12-31T22:59:59.999Z')
const Query8 = OData.Where<IEmployee>().DateField("Created").IsBetween(new Date(2023, 0, 1, 0, 0, 0, 0), new Date(2023, 11, 31, 23, 59, 59, 999)).ToString();

console.log(Query1);

console.log(Query2);

console.log(Query3);

console.log(Query4);

console.log(Query5);

console.log(Query6);

console.log(Query7);

console.log(Query8);
