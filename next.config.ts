import type { NextConfig } from "next";

// const isDevelopment = process.env.NODE_ENV === "development";
// const scriptSrc = isDevelopment
//   ? "script-src 'self' 'unsafe-inline' 'unsafe-eval'"
//   : "script-src 'self' 'unsafe-inline'";

const nextConfig: NextConfig = {
  // async headers() {
  //   return [
  //     {
  //       source: "/:path*",
  //       headers: [
  //         {
  //           key: "Content-Security-Policy",
  //           value: [
  //             "default-src 'self'",
  //             "base-uri 'self'",
  //             "connect-src 'self'",
  //             "font-src 'self' data: https://fonts.gstatic.com",
  //             "form-action 'self'",
  //             "frame-ancestors 'none'",
  //             "img-src 'self' blob: data:",
  //             "object-src 'none'",
  //             scriptSrc,
  //             "style-src 'self' 'unsafe-inline'",
  //             "style-src-elem 'self' 'unsafe-inline' https://fonts.googleapis.com",
  //             "worker-src 'self' blob:",
  //           ].join("; "),
  //         },  
  //         {
  //           key: "Referrer-Policy",
  //           value: "no-referrer",
  //         },
  //         {
  //           key: "X-Content-Type-Options",
  //           value: "nosniff",
  //         },
  //         {
  //           key: "Permissions-Policy",
  //           value:
  //             "camera=(), microphone=(), geolocation=(), payment=(), usb=(), serial=(), bluetooth=()",
  //         },
  //         {
  //           key: "X-Frame-Options",
  //           value: "DENY",
  //         },
  //       ],
  //     },
  //   ];
  // },
};

export default nextConfig;
