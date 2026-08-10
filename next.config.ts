import type { NextConfig } from "next";

const nextConfig: NextConfig = {
  allowedDevOrigins: ["**.localhost"],
  async headers() {
    return [
      {
        source: "/:path*",
        headers: [
          {
            key: "Origin-Agent-Cluster",
            value: "?1",
          },
        ],
      },
    ];
  },
  webpack: (config, { isServer }) => {


    // 减少文件监听，避免 EMFILE 错误
    config.watchOptions = {
      ...config.watchOptions,
      ignored: [
        '**/node_modules/**',
        '**/.git/**',
        '**/.next/**',
        '**/dist/**',
        '**/build/**',
      ],
      aggregateTimeout: 300,
      poll: 1000, // 使用轮询代替文件监听
    };

    return config;
  },
 
};

export default nextConfig;
