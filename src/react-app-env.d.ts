/// <reference types="vite/client" />

interface ImportMetaEnv {
  readonly VITE_AAD_CLIENT_ID: string
  readonly VITE_AAD_TENANT_ID: string
  readonly VITE_AAD_AUTHORITY: string
  readonly VITE_SHAREPOINT_SITE_ID: string
  readonly VITE_SHAREPOINT_PERSONNEL_LIST_ID: string
  readonly VITE_API_BASE_URL: string
}

interface ImportMeta {
  readonly env: ImportMetaEnv
}
