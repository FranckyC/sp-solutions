# SharePoint List Webhook Handler Function

## Description

This function allows you to handle updates on a list using the SharePoint. This function is written using TypeScript and has been implemented using the following [boilerplate](https://github.com/aequos-solutions/sp-solutions/tree/master/boilerplates/typescript-azure-function). 

### Why an azure functioon instead of a Microsoft Flow or Logic App trigger?

It has been designed mainly to replace an exiting Microsoft Flow or Logic App pulling trigger to save costs and get udpdates has they arrive in a SharePoint list (for instance in a site provisioning solution).

## Usage 

In the `webhookHandler.ts` file, set the event(s) you want to monitor in a list by setting the change request:

```
// Get changes for 'Add' operations only
const changeRequest: IChangeQuery = {
    query: {
        Add: true,
        Item: true,
        RecursiveAll: true,
        ChangeTokenStart: changeToken
    }
};
```

### Create the Webhook subscription

To first create the webhook subscription from the SharePoint list, you can either:
- Run the function for a first time, without any parameter. The subscription will be created wautomatically.
- Use PnP PowerShell, to create the subscription manually.

### Set the renewal period

By default, this value is set to 180 days. You can change it by setting the value in the `expirationRenewalDays` proeprty in the `SharePointWebHookhandler` instance:

```

```

## How to debug this function locally ?

- In VSCode, open the root folder `./`.
- Install all dependencies using `npm i`.
- Install [Azure CLI](https://docs.microsoft.com/en-us/cli/azure/install-azure-cli-windows?view=azure-cli-latest) on youre machine.
- Install Azure Function Core tools globaly using `npm install -g azure-functions-core-tools@2` (version 2).
- Install [DotNet Core SDK](https://dotnet.microsoft.com/download)
- In a Node.js console, build the solution using `npm run build:dev` cmd. For production use, execute `npm run build` (minified version of the JS code).
- In a Node.js console, from the `./dist` folder, run the following command `func start`.
- From VSCode, Launch the *'Debug Local Azure Function'* debug configuration 
- Send your requests either using Postman with the localhost address according to your settings (i.e. `http://localhost:7071/api/demoFunction`)
- Enjoy ;)

<p align="center"><img width="500px" src="./images/func_debug.png"/><p>

### Debug tests

- Set breakpoints directly in your **'spec.ts'** test files
- In VSCode, launch the *'Debug Jest all tests'* debug configuration
- In a Node.js console, build the solution using `npm run test`
- Enjoy ;)

### Azure Function Proxy configuration ###

This solution uses an Azure function proxy to get an only single endpoint URL for multiple functions. See the **proxies.json** file to see defined routes.

### Certificate generation ###

This solution uses an Azure AD application certificate connection to communicate with SharePoint. It means you need to generate the certificate private key in _*.pem_ format and copy it to the _'./config'_ folder of your function. You can create a certificate for your Azure AD app using this [procedure](https://docs.microsoft.com/en-us/sharepoint/dev/solution-guidance/security-apponly-azuread). Then use the [OpenSSL](https://wiki.openssl.org/index.php/Binaries) tool and the following command to generate the _.pem_ key from the _.pfx_ certificate:

```
openssl pkcs12 -in C:\<your_certificate>.pfx -nocerts -out C:\<your_certificate>.pem -nodes
```

You can also use the [`Get-PnPAzureCertificate`](https://docs.microsoft.com/en-us/powershell/module/sharepoint-pnp/get-pnpazurecertificate?view=sharepoint-ps) cmdlet to do the same:

```
Get-PnPAzureCertificate -CertificatePath "C:\<your_certificate>.pfx -" -CertificatePassword (ConvertTo-SecureString -String '<your_password>' -AsPlainText -Force)
```

## How to deploy the solution to Azure ? ##

### Development scenario

We recommend to use Visual Studio Code to work with this solution.

- In VSCode, download the [Azure Function](https://code.visualstudio.com/tutorials/functions-extension/getting-started) extension
- Sign-in to to Azure account into the extension
- In a Node.js console, build the application using the command `npm run build` (minified version)
- Use the **"Deploy to Function App"** feature (in the extension top bar) using the *'dist'* folder. Make sure you've run the `npm run build` cmd before.
- Upload the application settings file according to your environment (`local.settings.json`)

### Production scenario with CI

A `deploy.ps1` script is available to also deploy this function into your Azure environment.

- From you Azure portal, create a new empty function
- Set the `Azure_Function_Name` value in the `local.settings.json` accordingly.
- Login to Azure using `az login` then run `deploy.ps1` script with your parameters.
- Don't hesitate to refactor this to adapt to your situation.

***In both scenarios, you can test your function using Postman. If you test it using a SPFx component, don't forget to add the SharePoint domain to the CORS settings to allow this origin:***

<p align="center"><img width="500px" src="./images/cors_settings.png"/><p>




