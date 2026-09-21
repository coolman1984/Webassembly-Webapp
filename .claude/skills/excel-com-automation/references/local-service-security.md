# Securing a "runs on localhost, opens in your browser" tool

A tool that binds to `127.0.0.1` and opens itself in the default browser is not automatically safe just because
it's not reachable from the internet. **Any webpage the user has open in the same browser, in any tab, can still
try to talk to it** - that's the actual threat model for a local HTTP service, and it needs real defences, not
"it's just localhost."

## Bind to loopback only

```python
ThreadingHTTPServer(("127.0.0.1", port), Handler)
```

Never `0.0.0.0` - that would make the service reachable from other machines on the same network too.

## Validate the `Host` header on every request (DNS rebinding)

A malicious page can register a domain whose DNS record first resolves to a real server (so the browser loads
the page) and then, after the page's JavaScript starts making "same-origin" requests, gets its DNS answer
changed to `127.0.0.1`. The browser's same-origin check operates on hostname, not resolved IP, so this bypasses
it entirely - and it does not require an `Origin` header at all for a simple `GET`. The one thing that reliably
differs is the `Host` header the request carries, which will be the attacker's domain, not `127.0.0.1:<port>` or
`localhost:<port>`.

```python
def host_ok(self):
    return (self.headers.get("Host") or "").lower() in (f"127.0.0.1:{PORT}", f"localhost:{PORT}")
```

Check this **before** anything else, on every route, including plain `GET`s that only read data.

## Require a custom header (and check `Origin`) on every state-changing request

A cross-site `<form>` POST (or an `img`/`script` tag abusing a GET with side effects) can't set arbitrary custom
request headers - only `fetch`/`XMLHttpRequest` from a page that's deliberately trying to can. Require one on
every `PUT`/`POST`:

```python
def write_ok(self):
    origin = self.headers.get("Origin")
    return (origin is None or origin in ALLOWED_ORIGINS) and self.headers.get("X-Requested-With") == "your-app-name"
```

This, combined with the `Host` check above, closes both the DNS-rebinding read path and the simple cross-site
write path. Neither check alone is sufficient - `Host` catches what `Origin` can miss (a same-origin-looking GET
after rebinding carries no `Origin` header at all), and the custom-header/`Origin` pair catches what a bare
`Host` check doesn't (a legitimate-looking same-origin form submission from a *different* malicious page that
also happens to be served from `127.0.0.1` in some edge case).

## Never let a client hold a server thread open

Use a threaded server (`ThreadingHTTPServer`, daemon threads) with a socket-level `timeout` on the handler class,
and validate `Content-Length` before reading a request body (reject missing/negative/absurdly large values
outright) - an abandoned or malicious slow client should not be able to starve the service.

## Sanitise uploaded filenames hard

Strip path separators and drive letters, `..`, control characters, and cap length before using any part of a
client-supplied filename to build a path on disk. Test with path-traversal attempts, very long names, and
non-ASCII (including right-to-left script) names explicitly - not just "normal-looking" filenames.

## Neutralise CSV/spreadsheet formula injection on export

Any value you let a user later re-open in Excel (a CSV export, etc.) that could start with `=`, `+`, `@`, tab, or
CR must be prefixed (e.g. with a leading `'`) before being written - otherwise data that merely passed through
your system (never validated as "safe," just stored and later exported) can execute as a formula in the
recipient's spreadsheet application.

## Some corporate security tools break browser file uploads - have a fallback

An endpoint-security/DRM agent can intercept the browser's file-picker read and make the *original* file
unreadable to the browser (rewritten to an opaque temporary name, or the read simply throws) while the file
remains completely normal on disk to any trusted local process. If your tool's only way to get a file in is
"the browser reads it and uploads the bytes," this is a real, un-debuggable-from-your-side failure mode for a
meaningful fraction of corporate machines. Offer a second path - e.g. a designated folder the service reads
files from directly on disk when the user clicks "process" with nothing attached in the browser - so the tool
still works without needing to identify or work around whichever specific security product is interfering.

## Prefer atomic writes and versioned schema for any local database

Write a new SQLite file to a temp path in the same directory, then `os.replace()` it over the real path -
readers never see a half-written file, and a crash mid-write leaves the previous good file untouched. Stamp a
schema version in the database's own metadata and refuse to read a mismatched one (treat it as "no data yet"
rather than crashing) so a future version of the tool doesn't misinterpret an old file's layout.
