import { createRoot } from "react-dom/client";

import { App } from "./App";
import { Bridge, api, hostTransport } from "./bridge";

// `yarn dev` in a browser, or with ?mock in the app, uses made-up data (src/mock.ts)
const mock = import.meta.env.DEV && (!hostTransport() || new URLSearchParams(location.search).has("mock"));
const transport = mock ? (await import("./mock")).mockTransport() : hostTransport();
const root = createRoot(document.getElementById("root")!);

if (transport) {
  root.render(<App api={api(new Bridge(transport))} />);
} else {
  root.render(
    <p style={{ padding: 32 }}>
      This is the user interface of the Scriptless SharePoint app for Windows. Open it in the app.
    </p>,
  );
}
