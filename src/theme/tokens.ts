export const dashboardTokens = {
  cardRadius: "16px",
  cardBorder: "1px solid rgba(255,255,255,0.08)",
  cardBg: "rgba(16,16,20,0.75)",
  textPrimary: "#f5f7ff",
  textSecondary: "rgba(226,232,240,0.72)",
  accent: "#60a5fa",
  spacing: {
    xs: 4,
    sm: 8,
    md: 12,
    lg: 16,
    xl: 24,
  },
} as const;

export type DashboardTokens = typeof dashboardTokens;
