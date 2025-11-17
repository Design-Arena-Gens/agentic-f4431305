/** @type {import('next').NextConfig} */
const nextConfig = {
  reactStrictMode: true,
  experimental: {
    serverActions: {
      allowedOrigins: ["https://agentic-f4431305.vercel.app", "http://localhost:3000"],
    }
  }
};

module.exports = nextConfig;
