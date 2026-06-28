import { NextResponse } from "next/server";
import { cookies } from "next/headers";
import { getUsers, signSession } from "@/utils/auth";

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
