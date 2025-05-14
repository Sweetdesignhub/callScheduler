require("dotenv").config();
const express = require("express");
const { Client } = require("@microsoft/microsoft-graph-client");
const { ClientSecretCredential } = require("@azure/identity");
const crypto = require("crypto");
const cors = require("cors");
const morgan = require("morgan");

// Allow CORS for everything
require("isomorphic-fetch");

const app = express();
const PORT = 3000;
app.use(cors());
app.use(morgan("combined"));

app.use(express.json());

const credential = new ClientSecretCredential(
  process.env.TENANT_ID,
  process.env.CLIENT_ID,
  process.env.CLIENT_SECRET
);

// STEP 1: Initialize Graph Client
const getGraphClient = async () => {
  try {
    console.log("🔐 Getting Microsoft Graph access token...");
    const token = await credential.getToken([
      "https://graph.microsoft.com/.default",
    ]);
    console.log("✅ Access token acquired.");
    return Client.init({
      authProvider: (done) => done(null, token.token),
    });
  } catch (err) {
    console.error("❌ Failed to get token:", err);
    throw err;
  }
};

// STEP 2: Schedule route
app.post("/schedule", async (req, res) => {
  const { subject, duration, preferredWindow = {} } = req.body;

  console.log("📥 Received schedule request:");
  console.log({ subject, duration, preferredWindow });

  const now = new Date();
  const startTime =
    preferredWindow.start ||
    new Date(now.getTime() + 30 * 60 * 1000).toISOString();
  const endTime =
    preferredWindow.end ||
    new Date(now.getTime() + 6 * 60 * 60 * 1000).toISOString();

  try {
    const graphClient = await getGraphClient();

    // STEP 3: Check manager availability
    console.log(
      `📆 Checking manager (${process.env.MANAGER_EMAIL}) availability from ${startTime} to ${endTime}...`
    );

    const availability = await graphClient
      .api(`/users/${process.env.MANAGER_EMAIL}/calendar/getSchedule`)
      .post({
        schedules: [process.env.MANAGER_EMAIL],
        startTime: { dateTime: startTime, timeZone: process.env.TIME_ZONE },
        endTime: { dateTime: endTime, timeZone: process.env.TIME_ZONE },
        availabilityViewInterval: duration,
      });
    console.log("availablibity:", availability);
    const schedule = availability.value[0];
    console.log("schedule:", schedule);

    const freeSlot = schedule.scheduleItems.find(
      (item) => item.status === "free"
    );

    if (!freeSlot) {
      console.log("⚠️ No available time slots found.");
      return res
        .status(404)
        .json({ message: "No free slots available in preferred window." });
    }

    const meetingStart = freeSlot.start.dateTime;
    const meetingEnd = new Date(
      new Date(meetingStart).getTime() + duration * 60000
    ).toISOString();

    console.log("✅ Found free slot:");
    console.log(`⏱ Start: ${meetingStart}`);
    console.log(`⏱ End:   ${meetingEnd}`);

    // STEP 4: Create Teams meeting
    const event = {
      subject,
      start: { dateTime: meetingStart, timeZone: process.env.TIME_ZONE },
      end: { dateTime: meetingEnd, timeZone: process.env.TIME_ZONE },
      attendees: [
        {
          emailAddress: { address: process.env.MANAGER_EMAIL },
          type: "required",
        },
      ],
      isOnlineMeeting: true,
      onlineMeetingProvider: "teamsForBusiness",
    };

    console.log("📤 Creating Teams meeting via Graph API...");
    const createdEvent = await graphClient
      .api(`/users/${process.env.MANAGER_EMAIL}/events`)
      .post(event);

    console.log("✅ Meeting created successfully!");
    console.log(`🔗 Join URL: ${createdEvent.onlineMeeting?.joinUrl}`);

    res.json({
      message: "Meeting scheduled!",
      joinUrl: createdEvent.onlineMeeting?.joinUrl || "No join URL found",
      event: createdEvent,
    });
  } catch (error) {
    console.error("❌ Error during scheduling process:", error);
    res.status(500).json({ message: "Error scheduling meeting", error });
  }
});

app.post("/schedule-outlook", async (req, res) => {
  const { subject, startTime, endTime, attendees } = req.body;
  console.log("📥 Creating Outlook meeting:", req.body);

  try {
    // const token = await getAccessToken();
    // const client = Client.init({
    //   authProvider: (done) => done(null, token),
    // });
    const client = await getGraphClient();
    const event = {
      subject,
      start: {
        dateTime: startTime,
        timeZone: process.env.TIME_ZONE || "India Standard Time",
      },
      end: {
        dateTime: endTime,
        timeZone: process.env.TIME_ZONE || "India Standard Time",
      },
      attendees: attendees.map((email) => ({
        emailAddress: { address: email, name: email.split("@")[0] },
        type: "required",
      })),
      isOnlineMeeting: true,
      onlineMeetingProvider: "teamsForBusiness",
    };

    console.log("Event:", event);

    const createdEvent = await client
      .api(`/users/${process.env.MANAGER_EMAIL}/events`)
      .post(event);

    console.log("Created Event:", createdEvent);

    res.status(200).json({
      message: "✅ Meeting scheduled in Outlook with Teams link.",
      joinUrl: createdEvent.onlineMeeting?.joinUrl,
      eventId: createdEvent.id,
    });
  } catch (error) {
    console.error("❌ Error creating meeting:", error);
    res.status(500).json({ error: error.message || "Internal Server Error" });
  }
});

app.post("/auto-schedule-outlook", async (req, res) => {
  const { subject, durationMinutes, attendees, description } = req.body;
  const timeZone = process.env.TIME_ZONE || "India Standard Time";
  const managerEmail = process.env.MANAGER_EMAIL;

  try {
    const client = await getGraphClient();

    // Step 1: Find manager's nearest availability
    const now = new Date();
    const later = new Date(now.getTime() + 3 * 24 * 60 * 60 * 1000); // 3 days

    const findTimePayload = {
      attendees: [
        {
          type: "required",
          emailAddress: { address: managerEmail },
        },
      ],
      timeConstraint: {
        timeslots: [
          {
            start: { dateTime: now.toISOString(), timeZone },
            end: { dateTime: later.toISOString(), timeZone },
          },
        ],
      },
      meetingDuration: `PT${durationMinutes}M`,
      maxCandidates: 5,
      isOrganizerOptional: false,
      returnSuggestionReasons: true,
      minimumAttendeePercentage: 100,
    };

    const meetingTimeResults = await client
      .api(`/users/${managerEmail}/findMeetingTimes`)
      .post(findTimePayload);

    const bestSlot = meetingTimeResults.meetingTimeSuggestions?.[0];
    if (!bestSlot) {
      return res
        .status(400)
        .json({ message: "❌ No available time slot found." });
    }

    const { start, end } = bestSlot.meetingTimeSlot;

    // Step 2: Create the event in manager's calendar
    const event = {
      subject,
      body: {
        contentType: "HTML",
        content: description || "Meeting scheduled via auto-scheduler.",
      },
      start,
      end,
      attendees: [
        ...attendees.map((email) => ({
          emailAddress: { address: email },
          type: "required",
        })),
        {
          emailAddress: { address: managerEmail },
          type: "required",
        },
      ],
      isOnlineMeeting: true,
      onlineMeetingProvider: "teamsForBusiness",
    };

    const createdEvent = await client
      .api(`/users/${managerEmail}/events`)
      .post(event);

    // Step 3: Send email to attendees
    const emailMessage = {
      message: {
        subject: `📅 Scheduled Meeting: ${subject}`,
        body: {
          contentType: "HTML",
          content: `
              <p>Hi,</p>
              <p>You have been invited to a meeting scheduled by the manager.</p>
              <p><strong>Subject:</strong> ${subject}</p>
              <p><strong>Description:</strong> ${description}</p>
              <p><strong>Start:</strong> ${start.dateTime}</p>
              <p><strong>End:</strong> ${end.dateTime}</p>
              <p><strong>Join URL:</strong> <a href="${createdEvent.onlineMeeting?.joinUrl}">Click here</a></p>
              <p>Thanks.</p>
            `,
        },
        toRecipients: attendees.map((email) => ({
          emailAddress: { address: email },
        })),
      },
      saveToSentItems: true,
    };

    await client.api(`/users/${managerEmail}/sendMail`).post(emailMessage);

    res.status(200).json({
      message: "✅ Meeting scheduled and email sent.",
      joinUrl: createdEvent.onlineMeeting?.joinUrl,
      start: start.dateTime,
      end: end.dateTime,
      eventId: createdEvent.id,
    });
  } catch (err) {
    console.error("❌ Scheduling failed:", err);
    res.status(500).json({ error: err.message });
  }
});

app.listen(PORT, () => {
  console.log(`🚀 Server running at http://localhost:${PORT}`);
});
