import type { NextConfig } from "next";

const isGitHubPages = process.env.GITHUB_PAGES === "true";

const nextConfig: NextConfig = {
  output: isGitHubPages ? "export" : undefined,
  // Keep the app route at `/` during prerendering. Vite applies the repository
  // path to generated asset URLs for project Pages sites.
  basePath: "",
  assetPrefix: "",
  trailingSlash: isGitHubPages,
};

export default nextConfig;
