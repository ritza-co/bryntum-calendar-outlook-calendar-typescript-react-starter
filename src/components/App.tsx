import React from 'react';
import { MsalProvider } from '@azure/msal-react';
import { IPublicClientApplication } from '@azure/msal-browser';

import Calendar from './Calendar';
import '../css/App.css';
import ProvideAppContext from '../AppContext';

type AppProps = {
  pca: IPublicClientApplication
};

export default function App({ pca }: AppProps): React.JSX.Element {
    return (
        <MsalProvider instance={pca}>
            <ProvideAppContext>
                <Calendar />
            </ProvideAppContext>
        </MsalProvider>
    );
}