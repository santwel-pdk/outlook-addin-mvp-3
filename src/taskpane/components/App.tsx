import * as React from "react";
import Header from "./Header";
import EmailInfo from "./EmailInfo";
import SignalRStatus from "./SignalRStatus";
import { makeStyles } from "@fluentui/react-components";
import { useMailboxItem } from "../hooks/useMailboxItem";
import { useSignalR } from "../hooks/useSignalR";
import { getPlatformName } from "../utils/platform";
import { getBrowserEngine } from "../services/webview2Service";

interface AppProps {
  title: string;
}

const useStyles = makeStyles({
  root: {
    minHeight: "100vh",
  },
  platformInfo: {
    padding: "12px",
    textAlign: "center",
    fontSize: "12px",
    color: "#666",
  },
});

const App: React.FC<AppProps> = (props: AppProps) => {
  const styles = useStyles();
  const { emailData, isLoading, error } = useMailboxItem();
  const signalrState = useSignalR();
  //const platform = getPlatformName();
  //const browserEngine = getBrowserEngine();
  const dummyTitle = props.title;

  return (
    <div className={styles.root}>
      {/* <Header logo="assets/logo-filled.png" title={props.title} message="Welcome" />

      <div className={styles.platformInfo}>
        Platform: {platform} | Engine: {browserEngine}
      </div> */}

      <EmailInfo emailData={emailData} isLoading={isLoading} error={error} />
      
      <SignalRStatus signalrState={signalrState} />
    </div>
  );
};

export default App;
