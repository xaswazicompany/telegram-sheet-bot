"use client";

import { FormEvent, useState } from "react";

type SubmitState =
  | { status: "idle"; message: string }
  | { status: "loading"; message: string }
  | { status: "success"; message: string }
  | { status: "error"; message: string };

const initialState: SubmitState = { status: "idle", message: "" };

export default function Home() {
  const [state, setState] = useState<SubmitState>(initialState);

  async function handleSubmit(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    const form = event.currentTarget;
    const formData = new FormData(form);

    const payload = {
      fullName: String(formData.get("fullName") ?? "").trim(),
      email: String(formData.get("email") ?? "").trim(),
      company: String(formData.get("company") ?? "").trim(),
      projectType: String(formData.get("projectType") ?? "").trim(),
      budget: String(formData.get("budget") ?? "").trim(),
      message: String(formData.get("message") ?? "").trim(),
    };

    setState({ status: "loading", message: "Sending your request..." });

    try {
      const response = await fetch("/api/leads", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify(payload),
      });

      const result = (await response.json()) as { message?: string };

      if (!response.ok) {
        throw new Error(result.message ?? "Something went wrong.");
      }

      form.reset();
      setState({
        status: "success",
        message: result.message ?? "Your request was saved successfully.",
      });
    } catch (error) {
      setState({
        status: "error",
        message:
          error instanceof Error ? error.message : "Unable to submit the form.",
      });
    }
  }

  return (
    <main className="page-shell">
      <section className="hero-card">
        <div className="hero-copy">
          <p className="eyebrow">Professional Website + Google Sheets</p>
          <h1>Capture client requests on your website and store them securely in Google Sheets.</h1>
          <p className="lede">
            This starter uses a secure server-side API route, so your Google credentials never appear in the browser.
          </p>
          <div className="feature-grid">
            <div className="feature">
              <h2>Secure backend</h2>
              <p>Requests go from your form to a protected server route, then to Google Sheets.</p>
            </div>
            <div className="feature">
              <h2>Easy admin</h2>
              <p>Use Google Sheets as a lightweight CRM for leads, quotes, and customer messages.</p>
            </div>
            <div className="feature">
              <h2>Ready to deploy</h2>
              <p>Works well on platforms like Vercel once you add environment variables.</p>
            </div>
          </div>
        </div>

        <form className="lead-form" onSubmit={handleSubmit}>
          <div className="form-header">
            <p className="form-tag">Lead Form</p>
            <h2>Start your project</h2>
          </div>

          <label>
            Full name
            <input name="fullName" type="text" placeholder="John Carter" required />
          </label>

          <label>
            Email
            <input name="email" type="email" placeholder="john@company.com" required />
          </label>

          <label>
            Company
            <input name="company" type="text" placeholder="Acme Studio" />
          </label>

          <label>
            Project type
            <select name="projectType" defaultValue="Website development" required>
              <option>Website development</option>
              <option>E-commerce website</option>
              <option>Booking platform</option>
              <option>Landing page</option>
              <option>Internal dashboard</option>
            </select>
          </label>

          <label>
            Budget
            <select name="budget" defaultValue="$1,000 - $3,000" required>
              <option>$1,000 - $3,000</option>
              <option>$3,000 - $7,000</option>
              <option>$7,000 - $15,000</option>
              <option>$15,000+</option>
            </select>
          </label>

          <label>
            Project details
            <textarea
              name="message"
              rows={5}
              placeholder="Tell us what you want to build..."
              required
            />
          </label>

          <button type="submit" disabled={state.status === "loading"}>
            {state.status === "loading" ? "Sending..." : "Send Request"}
          </button>

          {state.message ? (
            <p className={`status ${state.status}`}>{state.message}</p>
          ) : null}
        </form>
      </section>
    </main>
  );
}

