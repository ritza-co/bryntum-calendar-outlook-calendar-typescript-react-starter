# How to connect and sync a React Bryntum Calendar to a Microsoft Outlook Calendar

This starter project was generated using the [Vite with TypeScript and React starter template](https://vite.dev/guide/#scaffolding-your-first-vite-project).

The code for the complete app is on the `completed-calendar` branch.

## Getting started

Install the dependencies by running the following command: 

```sh
npm install
```

Register your app in [Microsoft Entra admin center](https://entra.microsoft.com/) and make note of the Application (client) ID.

Create a  `.env` file in the root folder of your React Bryntum app and add the following variables to it:

```
VITE_MS_CLIENT_ID="your Application (client) ID"
VITE_MS_REDIRECT_URI="http://localhost:5173/"
```

## Running the app

Run the local dev server using the following command:

```sh
npm run dev
```

Open `http://localhost:5173` to see the Calendar app. 
