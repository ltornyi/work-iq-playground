import { InteractionStatus, type RedirectRequest } from "@azure/msal-browser";
import { useIsAuthenticated, useMsal } from "@azure/msal-react";
import {
  AuthProgress,
  HomeFooter,
  HomeHeader,
  HomeShell,
  SignInPanel,
  SignedInPanel,
} from "../components/home";

const loginRequest: RedirectRequest = {
  scopes: ["openid", "profile", "email"],
};

export default function Home() {
  const { instance, accounts, inProgress } = useMsal();
  const isAuthenticated = useIsAuthenticated();
  const account = accounts[0];

  const signIn = () => instance.loginRedirect(loginRequest);
  const signOut = () =>
    instance.logoutRedirect({
      account,
      postLogoutRedirectUri: window.location.origin,
    });

  return (
    <HomeShell>
      <HomeHeader />
      {inProgress !== InteractionStatus.None ? (
        <AuthProgress />
      ) : isAuthenticated && account ? (
        <SignedInPanel account={account} onSignOut={signOut} />
      ) : (
        <SignInPanel onSignIn={signIn} />
      )}
      <HomeFooter />
    </HomeShell>
  );
}
