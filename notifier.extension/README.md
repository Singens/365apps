# Notifier Extension
This repository contains the source code of the **Notifier Extension**.

**Notifier** is a SharePoint add-in that is available for installation through the Microsoft store for business apps AppSource. **Notifier Extension** is a free SharePoint Framework extension that depend on Notifier add-in.

![Notifier](./assets/images/notifier.preview.gif)

**Notifer** add-in provides functionality for displaying global messages to all pages in a SharePoint site. Due to security restriction of the store, the add-in cannot registry global script (ScriptLink) in the SharePoint site.
**Notifier Extension** implemented as SharePoint Framework Application Customizer can register global script and load it in all "modern" pages.  
Hereby, the **Notifier Extension** extend the **Notifier** functionality and both together provide end user complete solution. 

This approach is supported from the Microsoft after [the recent Store policy adjustment](https://dev.office.com/blogs/combining-store-add-ins-with-high-trust-permissions). 

# Using the repository

To build the package yourself, you'll need to clone and build the project.

Clone this repo by executing the following command in your console:

```
git clone https://github.com/Singens/365apps.git
```

Navigate to the cloned repository folder which should be the same as the repository name:

```
cd notifier.extension
```

Now run the following command to install the npm packages:

```
npm install
```

This will install the required npm packages and dependencies to build and run the client-side project.


Once the npm packages are installed, run the following command to start nodejs to host your extension and preview that in the SharePoint Online pages:

```
gulp serve
```

To build a deployment package use the following commands:
```
gulp bundle -ship
gulp package-solution -ship
```
Overview of the SharePoint Framework extensions you can find [here](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/extensions/overview-extensions). 