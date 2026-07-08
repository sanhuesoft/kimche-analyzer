import { NextResponse } from "next/server";
import { cookies } from "next/headers";
import { getUsers, signSession } from "@/utils/auth";

// Keep track of IP to set of logged in usernames
// Using globalThis to avoid reset during Next.js hot reloads in development
const globalForIpTrack = globalThis as unknown as {
  ipUserMap?: Map<string, Set<string>>;
};

const ipUserMap = globalForIpTrack.ipUserMap ?? new Map<string, Set<string>>();
if (process.env.NODE_ENV !== "production") {
  globalForIpTrack.ipUserMap = ipUserMap;
}

export async function POST(request: Request) {
  try {
    const { username, password } = await request.json();

    if (!username || !password) {
      return NextResponse.json(
        { error: "Usuario y contraseña son requeridos" },
        { status: 400 }
      );
    }

    const users = getUsers();
    const user = users.find(
      (u) => u.username.toLowerCase() === username.toLowerCase()
    );

    if (!user) {
      return NextResponse.json(
        { error: "Usuario o contraseña incorrectos" },
        { status: 401 }
      );
    }

    if (user.status !== "enabled") {
      return NextResponse.json(
        { error: "Este usuario ha sido deshabilitado" },
        { status: 403 }
      );
    }

    if (user.password !== password) {
      return NextResponse.json(
        { error: "Usuario o contraseña incorrectos" },
        { status: 401 }
      );
    }

    // IP detection logic has been disabled to prevent login issues.
    /*
    const forwardedFor = request.headers.get("x-forwarded-for");
    const ip = forwardedFor ? forwardedFor.split(",")[0].trim() : "127.0.0.1";

    const usernameKey = user.username.toLowerCase();
    if (!ipUserMap.has(ip)) {
      ipUserMap.set(ip, new Set());
    }
    const usersOnIp = ipUserMap.get(ip)!;
    usersOnIp.add(usernameKey);

    if (usersOnIp.size > 1) {
      console.log(
        `Dos usuarios iniciaron sesión de la misma ip. Usernames: ${Array.from(usersOnIp).join(", ")} (IP: ${ip})`
      );
      
      // La función que bloquea IPs al usar multi-cuenta está temporalmente desactivada:
      /*
      return NextResponse.json(
        { error: "Acceso bloqueado: detección de multi-cuenta para esta IP" },
        { status: 403 }
      );
      */
    /*
    }
    */

    const sessionToken = signSession(user.username);

    const cookieStore = await cookies();
    cookieStore.set("auth_session", sessionToken, {
      httpOnly: true,
      secure: process.env.NODE_ENV === "production",
      maxAge: 34560000, // 400 days (maximum browser duration)
      path: "/",
      sameSite: "lax",
    });

    return NextResponse.json({ success: true });
  } catch (error) {
    console.error("Login API error:", error);
    return NextResponse.json(
      { error: "Error interno del servidor" },
      { status: 500 }
    );
  }
}
