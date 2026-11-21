# Experimenting with Work IQ APIs

## Enable Work IQ API in your tenant

Run the command below as a Global Administrator in your tenant to create the Work IQ service principal in your organization. [Learn more](https://learn.microsoft.com/en-us/microsoft-365/copilot/extensibility/work-iq/enable-work-iq?tabs=azure-cli#enable-work-iq-api-in-your-organization).

```bash
export WORKIQ_APP_ID="fdcc1f02-fc51-4226-8753-f668596af7f7"
az ad sp create --id "$WORKIQ_APP_ID"
```

## Create a new app registration

Edit the `DISPLAY_NAME` value below to a name of your choice, then run the command to create a new app registration in your tenant. The command will return the `appId` of the newly created app registration, which you will need later.

```bash
export DISPLAY_NAME="Work IQ Playground Client"
# single-tenant
export SIGN_IN_AUDIENCE="AzureADMyOrg"   
az ad app create \
    --display-name "$DISPLAY_NAME" \
    --sign-in-audience "$SIGN_IN_AUDIENCE" \
    --is-fallback-public-client true \
    --query appId -o tsv
```

Create a service principal for the app registration you just created. Check the command output and make note of the `id` value. The `export SP_ID` command below will store the `id` value in an environment variable for later use.

```bash
export APP_ID="<appId from previous command>"
az ad sp create --id "$APP_ID"
export SP_ID=$(az ad sp show --id "$APP_ID" --query id -o tsv)
```

Add delegated permissions to the app registration. The command below will add the `WorkIQAgent.Ask` permission to the app registration. 

```bash
export WORKIQ_AGENT_ASK_SCOPE_ID="0b1715fd-f4bf-4c63-b16d-5be31f9847c2"
az ad app permission add --id "$APP_ID" --api "$WORKIQ_APP_ID" \
    --api-permissions "${WORKIQ_AGENT_ASK_SCOPE_ID}=Scope"
```

Make it effective and grant admin consent for the permission you just added. The command below will grant admin consent for the `WorkIQAgent.Ask` permission to the app registration.

```bash
az ad app permission grant --id "$SP_ID" --api "$WORKIQ_APP_ID" --scope "WorkIQAgent.Ask"
```

Check the setup in the Azure portal. Navigate to **Entra ID** > **App registrations** > *your app registration* > **Manage** >  **API permissions**. You should see the `WorkIQAgent.Ask` permission listed and admin consent granted.

## Add redirect URI to the app registration

We will run the app locally using Vite, which will serve the app on `http://localhost:5173`. We need to add this URL as a redirect URI to the app registration for the single-page application platform. The Azure CLI does not support it at this point - add the redirect URI in the Azure portal. Navigate to **Entra ID** > **App registrations** > *your app registration* > **Manage** >  **Authentication** > **Redirect URI configuration** . Add `http://localhost:5173/` as a redirect URI.

## Configure the SPA client

See details in the [README.md](workiq-client/README.md) file in the `workiq-client` directory.
