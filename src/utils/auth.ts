import crypto from "crypto";

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

export function getUsers(): User[] {
  try {
    let json = process.env.USERS_JSON || "[]";
    json = json.trim();
    if ((json.startsWith("'") && json.endsWith("'")) || (json.startsWith('"') && json.endsWith('"'))) {
      json = json.slice(1, -1);
    }
    return JSON.parse(json);
  } catch (e) {
    console.error("Error parsing USERS_JSON", e);
    return [];
  }
}
