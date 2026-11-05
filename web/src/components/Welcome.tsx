import { Body1, Button, Caption1, makeStyles, Title2, tokens } from "@fluentui/react-components";
import { AddRegular } from "@fluentui/react-icons";

const useStyles = makeStyles({
  root: {
    flexGrow: 1,
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    padding: "24px",
  },
  card: {
    maxWidth: "560px",
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    gap: "12px",
    textAlign: "center",
  },
  logo: { width: "80px", height: "80px", marginBottom: "8px" },
  text: { color: tokens.colorNeutralForeground2, textAlign: "center" },
  ways: { color: tokens.colorNeutralForeground3, marginTop: "8px" },
});

export function Welcome({ onAdd }: { onAdd(): void }) {
  const s = useStyles();
  return (
    <main className={s.root}>
      <div className={s.card}>
        <img src="./logo.png" alt="" className={s.logo} />
        <Title2 as="h2">Welcome to Scriptless SharePoint</Title2>
        <Body1 className={s.text}>
          Manage SharePoint Online without PowerShell. Start by adding a connection to your tenant.
        </Body1>
        <Button appearance="primary" size="large" icon={<AddRegular />} onClick={onAdd}>
          Add a connection
        </Button>
        <Caption1 className={s.ways}>
          Sign in with your Windows account, another account, an app with a certificate, or an access token.
        </Caption1>
      </div>
    </main>
  );
}
