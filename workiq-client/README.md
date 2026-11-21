# Work IQ client app

The SPA was created using the react-router template:

```bash
npx create-react-router@latest workiq-client
```

## Installation

Install the dependencies:

```bash
npm install
```

Generate types:

```bash
npm run typecheck
```

## Configuration

Create a `.env` file in the `workiq-client` directory and add the following environment variables. Replace the `VITE_ENTRA_CLIENT_ID` value with the `appId` of your app registration, and replace the `VITE_ENTRA_TENANT_ID` value with your tenant ID.

    # Application (client) ID from your Microsoft Entra app registration
    VITE_ENTRA_CLIENT_ID=<appId from previous command>
    # Directory (tenant) ID. Use your tenant GUID for a single-tenant app.
    VITE_ENTRA_TENANT_ID=<your tenant ID>

You can find you tenant ID by running the following command:

```bash
az account show --query tenantId -o tsv
```

You can list the app registrations in your tenant by running the following command:

```bash
az ad app list --query "[].{displayName: displayName, appId: appId}" -o table
```

## Development

Start the development server with HMR:

```bash
npm run dev
```

Your application will be available at `http://localhost:5173`.

## Building for Production

Create a production build:

```bash
npm run build
```

## Deployment

### Docker Deployment

To build and run using Docker:

```bash
docker build -t my-app .

# Run the container
docker run -p 3000:3000 my-app
```

The containerized application can be deployed to any platform that supports Docker, including:

- AWS ECS
- Google Cloud Run
- Azure Container Apps
- Digital Ocean App Platform
- Fly.io
- Railway

### DIY Deployment

If you're familiar with deploying Node applications, the built-in app server is production-ready.

Make sure to deploy the output of `npm run build`

```
├── package.json
├── package-lock.json (or pnpm-lock.yaml, or bun.lockb)
├── build/
│   ├── client/    # Static assets
│   └── server/    # Server-side code
```
