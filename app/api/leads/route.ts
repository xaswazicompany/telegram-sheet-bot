import { NextResponse } from "next/server";
import { appendLeadRow } from "@/lib/googleSheets";

type LeadPayload = {
  fullName: string;
  email: string;
  company?: string;
  projectType: string;
  budget: string;
  message: string;
};

function isValidEmail(email: string) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

function validateLead(payload: LeadPayload) {
  if (!payload.fullName || payload.fullName.length < 2) {
    return "Full name is required.";
  }

  if (!payload.email || !isValidEmail(payload.email)) {
    return "A valid email address is required.";
  }

  if (!payload.projectType) {
    return "Project type is required.";
  }

  if (!payload.budget) {
    return "Budget is required.";
  }

  if (!payload.message || payload.message.length < 10) {
    return "Please add at least 10 characters in the project details.";
  }

  return null;
}

export async function POST(request: Request) {
  try {
    const payload = (await request.json()) as LeadPayload;
    const validationError = validateLead(payload);

    if (validationError) {
      return NextResponse.json({ message: validationError }, { status: 400 });
    }

    await appendLeadRow([
      new Date().toISOString(),
      payload.fullName,
      payload.email,
      payload.company ?? "",
      payload.projectType,
      payload.budget,
      payload.message,
    ]);

    return NextResponse.json({
      message: "Thanks. Your request has been saved to Google Sheets.",
    });
  } catch (error) {
    console.error("Lead submission failed", error);

    return NextResponse.json(
      {
        message:
          "We could not save the request. Please check your Google Sheets configuration.",
      },
      { status: 500 },
    );
  }
}

