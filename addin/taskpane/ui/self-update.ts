/**
 * Reload the pane when the bridge serves a newer build. The id of the build this page loaded is read once; the
 * bridge's current id arrives with every handshake and from a poll, so a rebuild applies without restarting anything.
 */
export function createSelfUpdater(options: { isBusy(): boolean }) {
  const loadedBuildId = document.querySelector<HTMLMetaElement>('meta[name="addin-build-id"]')?.content ?? null;
  let reloadPending = false;
  let reloadScheduled = false;

  function reloadIfIdle(): void {
    if (!reloadPending || reloadScheduled || options.isBusy()) return;
    reloadScheduled = true;
    window.setTimeout(() => window.location.reload(), 400);
  }

  function check(serverBuildId: string | undefined): void {
    if (!serverBuildId || serverBuildId === "none") return;
    if (loadedBuildId === serverBuildId) return;
    reloadPending = true;
    reloadIfIdle();
  }

  function poll(): void {
    window.setInterval(() => {
      void fetch("/addin/build-id", { cache: "no-store" })
        .then((r) => r.json() as Promise<{ id: string }>)
        .then((j) => check(j.id))
        .catch(() => undefined);
    }, 20_000);
  }

  return { check, poll, reloadIfIdle };
}
