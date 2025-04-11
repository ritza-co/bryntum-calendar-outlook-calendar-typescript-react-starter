import React, {
    useContext,
    createContext,
    useState,
    useEffect,
    useMemo,
    useRef
} from 'react';
import { AuthCodeMSALBrowserAuthenticationProvider } from '@microsoft/microsoft-graph-client/authProviders/authCodeMsalBrowser';
import { InteractionType, PublicClientApplication } from '@azure/msal-browser';
import { useMsal } from '@azure/msal-react';

import { getUser } from './graphService';
import msalConfig from './msalConfig';

export interface AppUser {
  displayName?: string,
  email?: string,
  avatar?: string,
  timeZone?: string,
  timeFormat?: string
};

export interface AppError {
  message: string,
  debug?: string
};

type AppContext = {
  user?: AppUser;
  error?: AppError;
  signIn?: () => Promise<void>;
  signOut?: () => Promise<void>;
  displayError?: Function;
  clearError?: Function;
  authProvider?: AuthCodeMSALBrowserAuthenticationProvider;
  isLoading?: boolean;
}

interface ProvideAppContextProps {
  children: React.ReactNode;
}

const appContext = createContext<AppContext>({
    user         : undefined,
    error        : undefined,
    signIn       : undefined,
    signOut      : undefined,
    displayError : undefined,
    clearError   : undefined,
    authProvider : undefined,
    isLoading    : false
});

export function useAppContext(): AppContext {
    return useContext(appContext);
}

export default function ProvideAppContext({ children }: ProvideAppContextProps) {
    const auth = useProvideAppContext();
    return (
        <appContext.Provider value={auth}>
            {children}
        </appContext.Provider>
    );
}

function useProvideAppContext() {
    const msal = useMsal();
    const [user, setUser] = useState<AppUser | undefined>(undefined);
    const [error, setError] = useState<AppError | undefined>(undefined);
    const [isLoading, setIsLoading] = useState(true);
    const initialFetchDone = useRef(false);

    const displayError = (message: string, debug?: string) => {
        setError({ message, debug });
    };

    const clearError = () => {
        setError(undefined);
    };

    // Used by the Graph SDK to authenticate API calls
    const authProvider = useMemo(() => new AuthCodeMSALBrowserAuthenticationProvider(
        msal.instance as PublicClientApplication,
        {
            account         : msal.instance.getActiveAccount()!,
            scopes          : msalConfig.scopes,
            interactionType : InteractionType.Popup
        }
    ), [msal.instance]);

    useEffect(() => {
        const checkUser = async() => {
            if (!user) {
                if (initialFetchDone.current) return;
                try {
                    // Check if user is already signed in
                    const account = msal.instance.getActiveAccount();
                    if (account) {
                        initialFetchDone.current = true;
                        // Get the user from Microsoft Graph
                        const user = await getUser(authProvider);
                        setUser({
                            displayName : user.displayName || '',
                            email       : user.mail || user.userPrincipalName || '',
                            timeFormat  : user.mailboxSettings?.timeFormat || 'h:mm a',
                            timeZone    : user.mailboxSettings?.timeZone || 'UTC'
                        });
                    }
                }
                catch (err: any) {
                    displayError(err.message);
                }
                finally {
                    setIsLoading(false);
                }
            }
            else {
                setIsLoading(false);
            }
        };
        checkUser();
    }, [authProvider, msal.instance, user]);

    const signIn = async() => {
        setIsLoading(true);
        try {
            await msal.instance.loginPopup({
                scopes : msalConfig.scopes,
                prompt : 'select_account'
            });

            // Get the user from Microsoft Graph
            const user = await getUser(authProvider);

            setUser({
                displayName : user.displayName || '',
                email       : user.mail || user.userPrincipalName || '',
                timeFormat  : user.mailboxSettings?.timeFormat || '',
                timeZone    : user.mailboxSettings?.timeZone || 'UTC'
            });
        }
        catch (err: any) {
            displayError(err.message);
        }
        finally {
            setIsLoading(false);
        }
    };

    const signOut = async() => {
        setIsLoading(true);
        try {
            await msal.instance.logoutPopup();
            setUser(undefined);
        }
        finally {
            setIsLoading(false);
            window.location.reload();
        }
    };

    return {
        user,
        error,
        signIn,
        signOut,
        displayError,
        clearError,
        authProvider,
        isLoading
    };
}