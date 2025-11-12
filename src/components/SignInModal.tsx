import { useEffect, useState } from 'react';
import { FocusTrap } from 'focus-trap-react';
import { BryntumButton } from '@bryntum/calendar-react';
import { useAppContext } from '../AppContext';

function SignInModal() {
    const app = useAppContext();
    const [isVisible, setIsVisible] = useState(true);

    // Prevent page scrolling when modal is open
    useEffect(() => {
        if (isVisible) {
            const scrollY = window.scrollY;
            document.body.style.overflowY = 'hidden';
            window.scrollTo(0, 0);
            return () => {
                document.body.style.overflowY = 'auto';
                window.scrollTo(0, scrollY);
            };
        }
    }, [isVisible]);

    useEffect(() => {
        if (isVisible) {
            const handleEscape = (e: KeyboardEvent) => {
                if (e.key === 'Escape') {
                    setIsVisible(false);
                }
            };

            document.addEventListener('keydown', handleEscape);
            return () => {
                document.removeEventListener('keydown', handleEscape);
            };
        }
    }, [isVisible]);

    return (
        isVisible ? (
            <FocusTrap focusTrapOptions={{ initialFocus : '.b-button-tonal' }}>
                <div className="sign-in-modal">
                    <div className="sign-in-modal-content">
                        <div className="sign-in-modal-content-text">
                            <h2>Sign in with Microsoft</h2>
                            <p>Sign in to view and manage events from your Outlook Calendar</p>
                        </div>
                        <div className="close-modal">
                            <BryntumButton
                                icon='fa-times'
                                rendition="transparent"
                                onClick={() => setIsVisible(false)}
                            />
                        </div>
                        <BryntumButton
                            text={app.isLoading ? 'Signing in...' : 'Sign in with Microsoft'}
                            rendition="tonal"
                            onClick={() => app.signIn?.()}
                            disabled={app.isLoading}
                        />
                    </div>
                </div>
            </FocusTrap>
        ) : null
    );
}

export default SignInModal;