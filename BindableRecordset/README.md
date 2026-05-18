### Access MDB Database editor using a WinRT XAML ListView control bound to an ADODB Recordset

This project focuses on editing the **Customers** table from the classic **Northwind.mdb** database while bridging technologies several decades apart:

- a classic _ADODB Recordset_ is used to open the table
- a modern _XAML ListView_ control is bound to this Recordset with an implementation of the _IBindableVector_ interface

The binding is set to two-way mode and triggered on PropertyChanged so that every change is recorded in real time and saved to the database.

This whole process takes surprisingly little code to accomplish and each line is heavily commented for who is curious to take a look.

<img width="1282" height="747" alt="BindableRecordset" src="https://github.com/user-attachments/assets/efd36a2a-fcbc-4608-95f5-e242c9167b4c" />
