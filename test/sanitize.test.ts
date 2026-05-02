import { describe, it, expect } from "vitest";
import { htmlToText, sanitizeHtmlForDisplay, wrapUntrusted } from "../src/security/sanitize.js";

describe("htmlToText", () => {
  it("strips <script> blocks entirely", () => {
    const out = htmlToText("<p>hi</p><script>alert('x')</script><p>bye</p>");
    expect(out).not.toContain("alert");
    expect(out).toContain("hi");
    expect(out).toContain("bye");
  });

  it("strips <style> blocks", () => {
    const out = htmlToText("<style>.x{color:red}</style><div>hello</div>");
    expect(out).not.toContain("color");
    expect(out).toContain("hello");
  });

  it("removes inline event handlers", () => {
    const out = htmlToText('<a href="x" onclick="evil()">link</a>');
    expect(out).not.toContain("onclick");
    expect(out).not.toContain("evil()");
    expect(out).toContain("link");
  });

  it("removes images (tracking pixels)", () => {
    const out = htmlToText('hello <img src="https://tracker.example/pixel.gif" /> world');
    expect(out).not.toContain("tracker");
    expect(out).toContain("hello");
    expect(out).toContain("world");
  });

  it("decodes basic HTML entities", () => {
    expect(htmlToText("a &amp; b &lt;c&gt;")).toBe("a & b <c>");
    expect(htmlToText("&#65;")).toBe("A");
  });

  it("converts <br> and block-level closers to newlines", () => {
    const out = htmlToText("<p>one</p><p>two</p>three<br>four");
    expect(out.split(/\n+/)).toEqual(["one", "two", "three", "four"]);
  });
});

describe("sanitizeHtmlForDisplay", () => {
  it("removes scripts and on* handlers but keeps tags", () => {
    const out = sanitizeHtmlForDisplay('<p onclick="x">hi</p><script>bad()</script>');
    expect(out).toContain("<p");
    expect(out).toContain("hi");
    expect(out).not.toContain("onclick");
    expect(out).not.toContain("bad()");
  });
});

describe("wrapUntrusted", () => {
  it("includes the warning prefix and delimiters", () => {
    const wrapped = wrapUntrusted("body");
    expect(wrapped).toContain("untrusted");
    expect(wrapped).toContain("UNTRUSTED_EMAIL_CONTENT");
    expect(wrapped).toContain("body");
  });
});
