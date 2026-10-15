export interface ProviderLoginStatus {
  state: "idle" | "starting" | "waiting" | "success" | "cancelled" | "expired" | "error";
  userCode: string | null;
  verificationUri: string | null;
  expiresAt: number | null;
  message: string;
}
