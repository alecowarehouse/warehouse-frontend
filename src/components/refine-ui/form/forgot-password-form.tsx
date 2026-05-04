"use client";

import { cloneElement, isValidElement, useMemo, useState } from "react";
import { Loader2 } from "lucide-react";

import {
  useForgotPassword,
  useRefineOptions,
  useLink,
  useNotification,
} from "@refinedev/core";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import {
  Card,
  CardContent,
  CardFooter,
  CardHeader,
  CardTitle,
  CardDescription,
} from "@/components/ui/card";
import { Separator } from "@/components/ui/separator";
import { cn } from "@/lib/utils";

export const ForgotPasswordForm = () => {
  const [email, setEmail] = useState("");
  const { open } = useNotification();

  const Link = useLink();

  const { title } = useRefineOptions();
  const brandIcon = useMemo(() => {
    if (!isValidElement<{ style?: React.CSSProperties; className?: string }>(title.icon)) {
      return title.icon;
    }

    return cloneElement(title.icon, {
      style: {
        ...(title.icon.props.style || {}),
        width: "52px",
        height: "52px",
      },
      className: cn(title.icon.props.className, "h-13 w-13"),
    });
  }, [title.icon]);

  const { mutate: forgotPassword, isPending } = useForgotPassword();

  const handleForgotPassword = async (e: React.FormEvent<HTMLFormElement>) => {
    e.preventDefault();

    const normalizedEmail = email.trim();

    forgotPassword(
      {
        email: normalizedEmail,
      },
      {
        onSuccess: (result) => {
          if (result?.success) {
            open?.({
              type: "success",
              message: "Reset link sent",
              description:
                "If an account exists for this email, a password reset link has been sent.",
            });
            return;
          }

          open?.({
            type: "error",
            message: "Failed to send reset link",
            description:
              result?.error?.message || "Please check the email and try again.",
          });
        },
        onError: (error) => {
          open?.({
            type: "error",
            message: "Failed to send reset link",
            description:
              error?.message || "Please check the email and try again.",
          });
        },
      },
    );
  };

  return (
    <div
      className={cn(
        "flex",
        "flex-col",
        "items-center",
        "justify-center",
        "px-6",
        "py-8",
        "min-h-svh",
      )}
    >
      <div className={cn("flex", "items-center", "justify-center", "gap-3")}>
        {brandIcon && (
          <div className={cn("text-foreground", "shrink-0")}>{brandIcon}</div>
        )}
        <div className={cn("text-left")}>
          <p className={cn("text-xs", "uppercase", "tracking-[0.18em]", "text-muted-foreground")}>
            Aleco
          </p>
          <p className={cn("text-xl", "font-semibold", "leading-none", "text-foreground")}>
            Warehouse
          </p>
        </div>
      </div>

      <Card className={cn("sm:w-[456px]", "p-12", "mt-6")}>
        <CardHeader className={cn("px-0")}>
          <CardTitle
            className={cn(
              "text-blue-600",
              "dark:text-blue-400",
              "text-3xl",
              "font-semibold",
            )}
          >
            Forgot password
          </CardTitle>
          <CardDescription
            className={cn("text-muted-foreground", "font-medium")}
          >
            Enter your email to reset your password.
          </CardDescription>
        </CardHeader>

        <Separator />

        <CardContent className={cn("px-0")}>
          <form onSubmit={handleForgotPassword}>
            <div className={cn("flex", "flex-col", "gap-2")}>
              <Label htmlFor="email">Email</Label>
              <Input
                id="email"
                type="email"
                placeholder=""
                required
                value={email}
                onChange={(e) => setEmail(e.target.value)}
                disabled={isPending}
              />
            </div>
            <Button
              type="submit"
              size="lg"
              disabled={isPending}
              className={cn("w-full", "mt-6")}
            >
              {isPending ? (
                <span className={cn("inline-flex", "items-center", "gap-2")}>
                  <Loader2 className={cn("h-4", "w-4", "animate-spin")} />
                  Sending
                </span>
              ) : (
                "Send reset link"
              )}
            </Button>
          </form>

        </CardContent>
        <Separator />

        <CardFooter>
          <div className={cn("w-full", "text-center text-sm")}>
            <span className={cn("text-sm", "text-muted-foreground")}>
              Remember your password?{" "}
            </span>
            <Link
              to="/login"
              className={cn(
                "text-blue-600",
                "dark:text-blue-400",
                "font-semibold",
                "underline",
                "inline-flex",
                "items-center",
                "gap-2",
              )}
            >
              Sign in
            </Link>
          </div>
        </CardFooter>
      </Card>
    </div>
  );
};

ForgotPasswordForm.displayName = "ForgotPasswordForm";
