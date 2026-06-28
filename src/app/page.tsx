import { cookies } from "next/headers";
import DashboardClient from "./DashboardClient";
import Login from "./Login";
import { verifySession } from "@/utils/auth";

export default async function Page() {
  const cookieStore = await cookies();
  const authCookie = cookieStore.get("auth_session");
  const username = verifySession(authCookie?.value);

  if (!username) {
    return <Login />;
  }

  return <DashboardClient />;
}
