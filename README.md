# Microsoft Outlook Web Addin

Companies have explored multiple ways to boost their office productivity through add-ins (VSTO, COM, Web Add-ins). However, Microsoft has announced that they will be deprecating VSTO add-ins sometime in 2024.
As such, Web Add-ins have been the "Go To" alternative. Although less powerful than VSTO or COM add-ins, web add-ins have the advantage of integrating with both office desktop clients and web applications in just one codebase.
Further, the deployment model of Office add-ins are centralized and updated automatically (like with Office 365 enhancements).

As an example, I developed a simple Outlook Web Add-in below that integrates data from our CRM (Customer Relationship Management) and ERP (Enterprise Resource Planning) with Microsoft Outlook.
When an email is received, the add-in attempts to identify the sender from our CRM Accounts and gets their details and RFQ info, as well as connecting to our ERP, and fetching their billing data.

In Outlook Desktop Client:
![desktop](https://github.com/user-attachments/assets/619129b4-0781-4f16-9965-83bdd73981ef)

In Outlook Web App (OWA):
![owa](https://github.com/user-attachments/assets/134c80ca-91e9-45c0-8fba-bea97b1b5314)


Web add-ins are old news, however, since the deprecation of VSTO add-ins is imminent, there is no other choice but to develop future work productivity enhancements through web add-ins.
