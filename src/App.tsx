import { useEffect } from "react";
import { Authenticated, Refine, useGetIdentity } from "@refinedev/core";

import { RefineKbar, RefineKbarProvider } from "@refinedev/kbar";

import routerProvider, {
	NavigateToResource,
	DocumentTitleHandler,
	UnsavedChangesNotifier,
} from "@refinedev/react-router";
import { liveProvider } from "@refinedev/supabase";
import { BrowserRouter, Outlet, Route, Routes, useLocation, useNavigate } from "react-router";
import "./App.css";
import { Toaster } from "./components/refine-ui/notification/toaster";
import { useNotificationProvider } from "./components/refine-ui/notification/use-notification-provider";
import { ThemeProvider } from "./components/refine-ui/theme/theme-provider";
import { authProvider } from "./providers/auth";
import { dataProvider } from "./providers/data";
import { supabaseClient } from "./providers/supabase-client";
import { Home, Package, Users, ClipboardList, AlertTriangle } from "lucide-react";
import Dashboard from "./pages/dashboard";
import { Layout } from "./components/refine-ui/layout/layout";
import ItemList from "./pages/items/list";
import ItemCreate from "./pages/items/create";
import LoginPage from "./pages/login";
import RegisterPage from "./pages/register";
import ForgotPasswordPage from "./pages/forgot-password";
import UpdatePasswordPage from "./pages/update-password";
import UserList from "./pages/users/list";
import ProfilePage from "./pages/profile";
import ItemMovementListPage from "./pages/item-movements/list";
import ItemMovementCreatePage from "./pages/item-movements/create";
import BatchUploadMctPage from "./pages/item-movements/batch-upload";
import ItemMovementHistoryPage from "./pages/item-movements/history";
import EmergencyMovementListPage from "./pages/item-movements/emergencies-list";
import EmergencyMovementCreatePage from "./pages/item-movements/emergencies-create";
import EmergencyMovementHistoryPage from "./pages/item-movements/emergencies-history";

const isAdminRole = (role?: string | null) => (role ?? "").toLowerCase() === "admin";

const AdminRouteGuard = () => {
	const { data: identity, isLoading } = useGetIdentity<{ role?: string }>();

	if (isLoading) {
		return null;
	}

	if (!isAdminRole(identity?.role)) {
		return <NavigateToResource resource="dashboard" />;
	}

	return <Outlet />;
};

const AuthRecoveryRedirect = () => {
	const location = useLocation();
	const navigate = useNavigate();

	useEffect(() => {
		const hashParams = new URLSearchParams(location.hash.replace(/^#/, ""));
		const searchParams = new URLSearchParams(location.search);
		const recoveryType = hashParams.get("type") || searchParams.get("type");

		if (recoveryType === "recovery" && location.pathname !== "/update-password") {
			const suffix = location.hash || location.search;
			navigate(`/update-password${suffix}`, { replace: true });
		}
	}, [location.hash, location.pathname, location.search, navigate]);

	useEffect(() => {
		const { data } = supabaseClient.auth.onAuthStateChange((event) => {
			if (event === "PASSWORD_RECOVERY" && location.pathname !== "/update-password") {
				navigate(`/update-password${location.hash || location.search}`, { replace: true });
			}
		});

		return () => {
			data.subscription.unsubscribe();
		};
	}, [location.hash, location.pathname, location.search, navigate]);

	return null;
};

function App() {
	return (
		<BrowserRouter>
			<AuthRecoveryRedirect />
			<RefineKbarProvider>
				<ThemeProvider>
					<Refine
						dataProvider={dataProvider}
						liveProvider={liveProvider(supabaseClient)}
						authProvider={authProvider}
						routerProvider={routerProvider}
						notificationProvider={useNotificationProvider()}
						options={{
							syncWithLocation: true,
							warnWhenUnsavedChanges: true,
							projectId: "L6Lpfe-nuWoit-1MSid6",
							title: {
								text: "ALECO WAREHOUSE",
								icon: <img src="/aleco-icon.ico" alt="Logo" className="w-6 h-6" style={{ width: '24px', height: '24px' }} />
							}
						}}
						resources={[
							{
								name: 'dashboard',
								list: '/dashboard',
								meta: {
									label: 'Dashboard',
									icon: <Home className="w-4 h-4" />
								}
							},
							{
								name: 'items',
								list: '/items',
								create: '/items/create',
								meta: {
									label: 'Inventory',
									icon: <Package className="w-4 h-4" />
								}
							},
							{
								name: "mct",
								list: "/mct",
								create: "/mct/create",
								meta: {
									label: "MCT",
									icon: <ClipboardList className="w-4 h-4" />,
								},
							},
							{
								name: "emergencies",
								list: "/emergency",
								create: "/emergency/create",
								meta: {
									label: "Emergencies",
									icon: <AlertTriangle className="w-4 h-4" />,
								},
							},
							{
								name: 'users',
								list: '/users',
								meta: {
									label: 'Users',
									icon: <Users className="w-4 h-4" />
								},
							},
						]}
					>
						<Routes>
							{/* Auth routes */}
							<Route path="/update-password" element={<UpdatePasswordPage />} />

							<Route
								element={
									<Authenticated key="auth-pages" fallback={<Outlet />}>
										<NavigateToResource resource="dashboard" />
									</Authenticated>
								}>
								<Route path="/login" element={<LoginPage />} />
								<Route path="/register" element={<RegisterPage />} />
								<Route path="/forgot-password" element={<ForgotPasswordPage />} />
							</Route>

							<Route
								element={
									<Authenticated key="protected-routes">
										<Layout>
											<Outlet />
										</Layout>
									</Authenticated>
								}>
								<Route index element={<NavigateToResource resource="dashboard" />} />
								<Route path="/dashboard" element={<Dashboard />} />
								<Route path="/profile" element={<ProfilePage />} />
								<Route path="items">
									<Route index element={<ItemList />} />
									<Route path="create" element={<ItemCreate />} />
								</Route>
								<Route path="mct" element={<ItemMovementListPage />} />
								<Route path="mct/create" element={<ItemMovementCreatePage />} />
								<Route path="mct/batch-upload" element={<BatchUploadMctPage />} />
								<Route path="mct/history" element={<ItemMovementHistoryPage />} />
								<Route path="emergency" element={<EmergencyMovementListPage />} />
								<Route path="emergency/create" element={<EmergencyMovementCreatePage />} />
								<Route path="emergency/history" element={<EmergencyMovementHistoryPage />} />
								<Route path="item-movements/*" element={<NavigateToResource resource="mct" />} />
								<Route path="issue-return/*" element={<NavigateToResource resource="mct" />} />
								<Route path="users" element={<AdminRouteGuard />}>
									<Route index element={<UserList />} />
								</Route>

								{/* Resource routes will be handled by Refine */}
							</Route>
						</Routes>
						<Toaster />
						<RefineKbar />
						<UnsavedChangesNotifier />
						<DocumentTitleHandler
								handler={({ autoGeneratedTitle, pathname }) => {
									if (pathname === "/login") return "Sign in | Warehouse";
									if (pathname === "/register") return "Sign up | Warehouse";
									if (pathname === "/forgot-password") return "Forgot password | Warehouse";
									if (pathname === "/update-password") return "Update password | Warehouse";
									if (pathname === "/profile") return "Profile | Warehouse";
									if (pathname === "/mct") return "MCT | Warehouse";
									if (pathname === "/mct/create") return "Create MCT | Warehouse";
									if (pathname === "/mct/batch-upload") return "Batch Upload MCTs | Warehouse";
									if (pathname === "/mct/history") return "MCT History | Warehouse";
									if (pathname === "/emergency") return "Emergency | Warehouse";
									if (pathname === "/emergency/create") return "Create Emergency | Warehouse";
									if (pathname === "/emergency/history") return "Emergency History | Warehouse";

									if (autoGeneratedTitle.includes(" | ")) {
										return autoGeneratedTitle.replace(/\s\|\s.*$/, " | Warehouse");
									}
								return `${autoGeneratedTitle} | Warehouse`;
							}}
						/>
					</Refine>

				</ThemeProvider>
			</RefineKbarProvider>
		</BrowserRouter>
	);
}

export default App;
