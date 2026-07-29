import crypto from "crypto";
import fs from "fs";

export interface User {
  username: string;
  password?: string;
  status: "enabled" | "disabled";
}

const SECRET = process.env.AUTH_SECRET || "default-secret-key-replace-this-in-production";

export function signSession(username: string): string {
  const hash = crypto.createHmac("sha256", SECRET).update(username).digest("hex");
  return `${username}.${hash}`;
}

export function verifySession(sessionStr: string | undefined): string | null {
  if (!sessionStr) return null;
  const parts = sessionStr.split(".");
  if (parts.length !== 2) return null;
  const [username, hash] = parts;
  const expectedHash = crypto.createHmac("sha256", SECRET).update(username).digest("hex");
  if (hash === expectedHash) {
    const users = getUsers();
    const user = users.find(u => u.username.toLowerCase() === username.toLowerCase());
    if (user && user.status === "enabled") {
      return username;
    }
  }
  return null;
}

function parseUsersJson(rawJson: string): User[] {
  let cleaned = rawJson.trim();
  if ((cleaned.startsWith("'") && cleaned.endsWith("'")) || (cleaned.startsWith('"') && cleaned.endsWith('"'))) {
    cleaned = cleaned.slice(1, -1);
  }
  // Strip trailing commas or periods before ] or }
  cleaned = cleaned.replace(/[,.]\s*([\]}])/g, "$1");
  return JSON.parse(cleaned);
}

export function getUsers(): User[] {
  const usersFilePath = process.env.USERS_FILE_PATH || "/etc/kimche/users.json";

  // Check if mounted users file exists (remote deployment)
  try {
    if (fs.existsSync(usersFilePath)) {
      const content = fs.readFileSync(usersFilePath, "utf-8").trim();
      if (content) {
        return parseUsersJson(content);
      }
    }
  } catch (e) {
    console.error(`Error reading or parsing users file at ${usersFilePath}:`, e);
  }

  // Fallback to process.env.USERS_JSON (local development)
  try {
    const json = process.env.USERS_JSON || "[]";
    return parseUsersJson(json);
  } catch (e) {
    console.error("Error parsing USERS_JSON env variable:", e);
    return [];
  }
}

