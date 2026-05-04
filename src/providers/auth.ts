import type { AuthProvider } from "@refinedev/core";

import { setRememberSessionPreference, supabaseClient } from "./supabase-client";
import { SITE_URL } from "./constants";

const normalizeAppRole = (value?: string | null): "admin" | "user" => {
	if ((value ?? "").toLowerCase() === "admin") {
		return "admin";
	}

	return "user";
};

const getAppOrigin = () => {
	if (SITE_URL) {
		return SITE_URL;
	}

	if (typeof window !== "undefined") {
		return window.location.origin;
	}

	return undefined;
};

const getUpdatePasswordRedirectUrl = () => {
	const baseUrl = getAppOrigin();

	if (!baseUrl) {
		return undefined;
	}

	return `${baseUrl.replace(/\/$/, "")}/update-password`;
};

export const authProvider: AuthProvider = {
	login: async ({ email, password, providerName, rememberMe }) => {
		// sign in with oauth
		try {
			if (providerName) {
				setRememberSessionPreference(true);
				const { data, error } = await supabaseClient.auth.signInWithOAuth({
					provider: providerName,
				});

				if (error) {
					return {
						success: false,
						error,
					};
				}

				if (data?.url) {
					return {
						success: true,
					};
				}
			}

			// sign in with email and password
			setRememberSessionPreference(Boolean(rememberMe));
			const { data, error } = await supabaseClient.auth.signInWithPassword({
				email,
				password,
			});

			if (error) {
				const message =
					error.message?.toLowerCase().includes("invalid login credentials")
						? "Invalid email or password."
						: "Unable to sign in. Please try again.";
				return {
					success: false,
					error: {
						name: "Sign in failed",
						message,
					},
				};
			}

			if (data?.user) {
				return {
					success: true,
					redirectTo: "/dashboard",
				};
			}
		} catch (error: any) {
			return {
				success: false,
				error,
			};
		}

		return {
			success: false,
			error: {
				message: "Login failed",
				name: "Invalid email or password",
			},
		};
	},
	register: async (params) => {
		const { email, password, name, firstName, lastName } = (params ?? {}) as {
			email?: string;
			password?: string;
			name?: string;
			firstName?: string;
			lastName?: string;
		};
		const resolvedName =
			name?.trim() ||
			[firstName?.trim(), lastName?.trim()].filter(Boolean).join(" ") ||
			undefined;

		if (!email || !password) {
			return {
				success: false,
				error: {
					message: "Register failed",
					name: "Email and password are required",
				},
			};
		}

		try {
			const { data, error } = await supabaseClient.auth.signUp({
				email,
				password,
				options: {
					data: {
						name: resolvedName,
						first_name: firstName,
						last_name: lastName,
					},
				},
			});

			if (error) {
				return {
					success: false,
					error,
				};
			}

			if (data) {
				return {
					success: true,
				};
			}
		} catch (error: any) {
			return {
				success: false,
				error,
			};
		}

		return {
			success: false,
			error: {
				message: "Register failed",
				name: "Invalid email or password",
			},
		};
	},
	forgotPassword: async ({ email }) => {
		const normalizedEmail = String(email ?? "").trim();

		if (!normalizedEmail) {
			return {
				success: false,
				error: {
					message: "Enter the email address for your account.",
					name: "Email is required",
				},
			};
		}

		try {
			const redirectTo = getUpdatePasswordRedirectUrl();
			const { data, error } = await supabaseClient.auth.resetPasswordForEmail(
				normalizedEmail,
				redirectTo ? { redirectTo } : undefined,
			);

			if (error) {
				return {
					success: false,
					error,
				};
			}

			if (data) {
				return {
					success: true,
				};
			}
		} catch (error: any) {
			return {
				success: false,
				error,
			};
		}

		return {
			success: false,
			error: {
				message: "Forgot password failed",
				name: "Invalid email",
			},
		};
	},
	updatePassword: async ({ password }) => {
		try {
			const { data, error } = await supabaseClient.auth.updateUser({
				password,
			});

			if (error) {
				return {
					success: false,
					error,
				};
			}

			if (data) {
				return {
					success: true,
					redirectTo: "/",
				};
			}
		} catch (error: any) {
			return {
				success: false,
				error,
			};
		}
		return {
			success: false,
			error: {
				message: "Update password failed",
				name: "Invalid password",
			},
		};
	},
	logout: async () => {
		const { error } = await supabaseClient.auth.signOut();

		if (error) {
			return {
				success: false,
				error,
			};
		}

		return {
			success: true,
			redirectTo: "/login",
		};
	},
	onError: async (error) => {
		if (error?.code === "PGRST301" || error?.code === 401) {
			return {
				logout: true,
			};
		}

		return { error };
	},
	check: async () => {
		try {
			const { data } = await supabaseClient.auth.getSession();
			const { session } = data;

			if (!session) {
				return {
					authenticated: false,
					error: {
						message: "Check failed",
						name: "Session not found",
					},
					logout: true,
					redirectTo: "/login",
				};
			}
		} catch (error: any) {
			return {
				authenticated: false,
				error: error || {
					message: "Check failed",
					name: "Session not found",
				},
				logout: true,
				redirectTo: "/login",
			};
		}

		return {
			authenticated: true,
		};
	},
	getPermissions: async () => {
		const { data } = await supabaseClient.auth.getUser();
		const currentUser = data.user;

		if (!currentUser) {
			return null;
		}

		const metadata = currentUser.user_metadata as { role?: string } | undefined;
		const appMetadata = currentUser.app_metadata as { role?: string } | undefined;

		const { data: dbUser } = await supabaseClient
			.from("users")
			.select("role")
			.eq("id", currentUser.id)
			.maybeSingle();

		return normalizeAppRole(dbUser?.role ?? metadata?.role ?? appMetadata?.role);
	},
	getIdentity: async () => {
		const { data } = await supabaseClient.auth.getUser();

		if (data?.user) {
			const metadata = data.user.user_metadata as
				| { name?: string; first_name?: string; last_name?: string; role?: string }
				| undefined;
			const appMetadata = data.user.app_metadata as { role?: string } | undefined;

			const { data: dbUser } = await supabaseClient
				.from("users")
				.select("name,email,first_name,last_name,role,avatar_url")
				.eq("id", data.user.id)
				.maybeSingle();

			const firstName = dbUser?.first_name || metadata?.first_name || null;
			const lastName = dbUser?.last_name || metadata?.last_name || null;
			const fullName =
				dbUser?.name ||
				metadata?.name ||
				[firstName, lastName].filter(Boolean).join(" ") ||
				data.user.email;
			const avatarUrl = dbUser?.avatar_url ?? null;

			return {
				...data.user,
				name: fullName,
				email: dbUser?.email || data.user.email,
				first_name: firstName,
				last_name: lastName,
				firstName,
				lastName,
				fullName,
				avatar_url: avatarUrl,
				avatar: avatarUrl,
				role: normalizeAppRole(dbUser?.role ?? metadata?.role ?? appMetadata?.role),
			};
		}

		return null;
	},
};
