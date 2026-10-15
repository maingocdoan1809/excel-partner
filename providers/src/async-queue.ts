/**
 * Minimal push/pull queue that bridges callback-style CLI output into an
 * `AsyncIterable`. `end()` flushes remaining items and then completes;
 * `fail()` rejects the consumer after buffered items are drained.
 */
export class AsyncQueue<T> implements AsyncIterable<T> {
  private readonly items: T[] = [];
  private waiter: { resolve: () => void } | null = null;
  private done = false;
  private failure: unknown = null;

  push(item: T): void {
    if (this.done) return;
    this.items.push(item);
    this.wake();
  }

  end(): void {
    this.done = true;
    this.wake();
  }

  fail(error: unknown): void {
    if (this.done) return;
    this.failure = error ?? new Error("queue failed");
    this.done = true;
    this.wake();
  }

  get isDone(): boolean {
    return this.done;
  }

  private wake(): void {
    const waiter = this.waiter;
    this.waiter = null;
    waiter?.resolve();
  }

  async *[Symbol.asyncIterator](): AsyncIterator<T> {
    while (true) {
      if (this.items.length > 0) {
        yield this.items.shift() as T;
        continue;
      }
      if (this.done) {
        if (this.failure) throw this.failure;
        return;
      }
      await new Promise<void>((resolve) => {
        this.waiter = { resolve };
      });
    }
  }
}
