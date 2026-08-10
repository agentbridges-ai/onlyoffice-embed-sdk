import { createFileRoute } from "@tanstack/react-router";
import { HomePage } from "@/features/marketing";

export const Route = createFileRoute("/")({
  head: () => ({
    meta: [
      { title: "OnlyOffice Embed SDK — 浏览器端文档编辑器" },
      {
        name: "description",
        content:
          "将 OnlyOffice 集成到你的 Web 应用，无需 Document Server，支持 Word / Excel / PPT。",
      },
    ],
  }),
  component: HomePage,
});
