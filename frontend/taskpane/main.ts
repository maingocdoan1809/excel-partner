import { boot } from "./chat/boot.js";

// Excel, Word and PowerPoint share this pane; everything the pane does starts in chat/boot.ts.
Office.onReady((info) => boot(String(info.host)));
