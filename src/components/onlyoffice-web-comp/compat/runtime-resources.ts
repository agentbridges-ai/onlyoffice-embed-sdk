import {
  getStaticResource,
  registerOnlyOfficeStaticResource,
  resetOnlyOfficeStaticResource,
  type OnlyOfficeStaticResourceOptions,
} from "../const";
import { ONLYOFFICE_EMBED_SDK_VERSION } from "./version";

export type OfficeRuntimeResourceOperation =
  | "prepare-document"
  | "prefetch-recommended"
  | "load-all"
  | "check-health"
  | "download-font"
  | "install-font-preset"
  | "remove-font";

/** Readiness describes streamed entry points, not an offline asset install. */
export type OfficeRuntimeReadiness =
  | "checking"
  | "ready"
  | "needs-download"
  | "update-available"
  | "updating"
  | "paused"
  | "repair-needed"
  | "error";
export type OfficeDocumentResourceType = "word" | "cell" | "slide";
export type OfficeFontPreset = "basic" | "office-compatibility";
export type RuntimeCacheCategory =
  | "fonts"
  | "core"
  | OfficeDocumentResourceType;

export type RuntimeFontFamily = {
  id: string;
  name: string;
  bytes: number;
  paths: string[];
  downloaded: boolean;
  removable: boolean;
};

export type RuntimeCacheFailure = { path: string; reason: string };
export type RuntimeCacheCategoryProgress = {
  category: RuntimeCacheCategory;
  completedFiles: number;
  totalFiles: number;
  completedBytes: number;
  totalBytes: number;
};

export type RuntimeCacheProgress = {
  phase: "checking" | "ready" | "loading" | "complete" | "error";
  completedFiles: number;
  totalFiles: number;
  completedBytes: number;
  totalBytes: number;
  failedFiles: number;
  failures?: RuntimeCacheFailure[];
  categories: RuntimeCacheCategoryProgress[];
};

export type ResourceErrorCode =
  | "offline"
  | "network"
  | "timeout"
  | "integrity"
  | "quota"
  | "manifest"
  | "incompatible"
  | "storage"
  | "aborted";
export type ResourcePhase =
  | "idle"
  | "planning"
  | "downloading"
  | "verifying"
  | "activating"
  | "repairing"
  | "paused";
export type ResourceScope =
  | "recommended"
  | "document"
  | "all"
  | "repair"
  | "fonts";
export type ResourceProfile =
  | "base"
  | "word"
  | "cell"
  | "slide"
  | "fonts-basic"
  | "fonts-office-compat";
export type ResourcePlanRequest = {
  scope: ResourceScope;
  documentType?: OfficeDocumentResourceType;
  profiles?: ResourceProfile[];
};
export interface ResourcePlan {
  planId: string;
  releaseId: string;
  scope: ResourceScope;
  profiles: string[];
  totalBytes: number;
  downloadBytes: number;
  reusedBytes: number;
}
export interface FailedResource {
  path: string;
  code: ResourceErrorCode;
  attempts: number;
}

export interface RequiredReleaseIdentity {
  releaseId: string;
  manifestSha256: string;
  packageVersion: string;
  hostBuildId: string;
}

export interface ResourceInstallerSnapshot {
  installedRelease: string | null;
  targetRelease: string | null;
  availableRelease: string | null;
  availablePackageVersion: string | null;
  readiness: OfficeRuntimeReadiness;
  phase: ResourcePhase;
  storageMode: "cache-storage" | "http-cache";
  currentChunk: string | null;
  currentChunkIndex: number;
  currentChunkCount: number;
  downloadedBytes: number;
  downloadBytes: number;
  verifiedBytes: number;
  verifyBytes: number;
  bytesPerSecond: number;
  failedResources: FailedResource[];
  canPause: boolean;
  canResume: boolean;
  canRetry: boolean;
  errorCode: ResourceErrorCode | null;
  installedProfiles: ResourceProfile[];
}

export interface OfficeRuntimeResourceInstaller {
  plan(request: ResourcePlanRequest): Promise<ResourcePlan>;
  apply(plan: ResourcePlan): Promise<void>;
  checkForUpdates(): Promise<void>;
  checkHealth(options?: { deep?: boolean }): Promise<void>;
  repair(options: {
    scope: "required" | "installed" | "all";
  }): Promise<void>;
  pause(): void;
  resume(): Promise<void>;
  cancel(): void;
  getInstallerSnapshot(): ResourceInstallerSnapshot;
  subscribeInstaller(
    listener: (snapshot: ResourceInstallerSnapshot) => void,
  ): () => void;
  getInstalledPaths(): string[];
}

export type OfficeRuntimePackSnapshot = {
  id: "core" | "fonts" | OfficeDocumentResourceType;
  ready: boolean;
  completedBytes: number;
  totalBytes: number;
};

export type OfficeRuntimeResourceProgress = RuntimeCacheProgress;

export type OfficeRuntimeResourceSnapshot = {
  packageVersion: string;
  assetVersion: string;
  readiness: OfficeRuntimeReadiness;
  packs: OfficeRuntimePackSnapshot[];
  progress: RuntimeCacheProgress;
  fonts: RuntimeFontFamily[];
  verifiedFontPaths: string[];
  operation: OfficeRuntimeResourceOperation | null;
  error: { code: ResourceErrorCode; path?: string } | null;
  installedRelease: string | null;
  targetRelease: string | null;
  availableRelease: string | null;
  availablePackageVersion: string | null;
  /** This facade streams resources and never owns an offline release. */
  storageMode: "cache-storage" | "http-cache";
  phase: ResourcePhase;
  currentChunk: string | null;
  currentChunkIndex: number;
  currentChunkCount: number;
  downloadedBytes: number;
  downloadBytes: number;
  verifiedBytes: number;
  verifyBytes: number;
  bytesPerSecond: number;
  failedResources: FailedResource[];
  canPause: boolean;
  canResume: boolean;
  canRetry: boolean;
};

export type OfficeRuntimeResourceListener = (
  snapshot: OfficeRuntimeResourceSnapshot,
) => void;

export type OfficeRuntimeResourceManagerOptions = {
  storage?: Storage;
  fetch?: typeof fetch;
  cacheStorage?: CacheStorage;
  assetBaseUrl?: string | URL;
  releaseInstaller?: OfficeRuntimeResourceInstaller;
  canonicalOrigin?: string;
  allowLocalTestMode?: boolean;
  requiredReleaseIdentity?: RequiredReleaseIdentity;
  /** Native embed-sdk resource registration options. */
  cdnOrigin?: string | null;
  onlyofficeVersion?: string | null;
  assetManifestDigest?: string | null;
};

export class OfficeRuntimeResourceCompatibilityError extends Error {
  readonly operation: string;

  constructor(operation: string, message?: string) {
    super(
      message ||
        `${operation} requires an offline release manifest, which the direct-embed streaming resource facade does not provide`,
    );
    this.name = "OfficeRuntimeResourceCompatibilityError";
    this.operation = operation;
  }
}

const EMPTY_PROGRESS: OfficeRuntimeResourceProgress = {
  phase: "ready",
  completedFiles: 0,
  totalFiles: 0,
  completedBytes: 0,
  totalBytes: 0,
  failedFiles: 0,
  failures: [],
  categories: [],
};

/**
 * Compatibility facade for applications that used the browser package's
 * resource manager.
 *
 * embed-sdk serves immutable runtime files from registered HTTP(S) resources
 * and lets the browser HTTP cache stream them on demand. It does not copy the
 * full runtime into CacheStorage or implement per-font installation. Methods
 * that would imply those guarantees fail with a named compatibility error.
 */
export class OfficeRuntimeResourceManager {
  private readonly fetchImpl: typeof fetch;
  private readonly listeners = new Set<OfficeRuntimeResourceListener>();
  private lastHealthCheck = 0;
  private snapshot: OfficeRuntimeResourceSnapshot;

  private constructor(options: OfficeRuntimeResourceManagerOptions) {
    if (
      options.releaseInstaller ||
      options.canonicalOrigin !== undefined ||
      options.allowLocalTestMode !== undefined ||
      options.requiredReleaseIdentity
    ) {
      throw new OfficeRuntimeResourceCompatibilityError(
        "resource-manager-options",
        "releaseInstaller, canonicalOrigin, allowLocalTestMode, and requiredReleaseIdentity require onlyoffice-browser's transactional offline runtime",
      );
    }
    this.fetchImpl = options.fetch ?? globalThis.fetch.bind(globalThis);
    const resourceOptions = resolveStaticResourceOptions(options);
    if (resourceOptions) registerOnlyOfficeStaticResource(resourceOptions);
    const resource = getStaticResource();
    this.snapshot = {
      packageVersion: ONLYOFFICE_EMBED_SDK_VERSION,
      assetVersion: resource.version.onlyofficeSdk,
      readiness: "needs-download",
      packs: (["core", "fonts", "word", "cell", "slide"] as const).map(
        (id) => ({
          id,
          ready: false,
          completedBytes: 0,
          totalBytes: 0,
        }),
      ),
      progress: { ...EMPTY_PROGRESS },
      fonts: [],
      verifiedFontPaths: [],
      operation: null,
      error: null,
      installedRelease: null,
      targetRelease: null,
      availableRelease: null,
      availablePackageVersion: null,
      storageMode: "http-cache",
      phase: "idle",
      currentChunk: null,
      currentChunkIndex: 0,
      currentChunkCount: 0,
      downloadedBytes: 0,
      downloadBytes: 0,
      verifiedBytes: 0,
      verifyBytes: 0,
      bytesPerSecond: 0,
      failedResources: [],
      canPause: false,
      canResume: false,
      canRetry: false,
    };
  }

  static async create(
    options: OfficeRuntimeResourceManagerOptions = {},
  ): Promise<OfficeRuntimeResourceManager> {
    return new OfficeRuntimeResourceManager(options);
  }

  getSnapshot(): OfficeRuntimeResourceSnapshot {
    return {
      ...this.snapshot,
      packs: this.snapshot.packs.map((pack) => ({ ...pack })),
      progress: {
        ...this.snapshot.progress,
        failures: (this.snapshot.progress.failures ?? []).map((failure) => ({
          ...failure,
        })),
        categories: this.snapshot.progress.categories.map((category) => ({
          ...category,
        })),
      },
      verifiedFontPaths: [...this.snapshot.verifiedFontPaths],
      fonts: this.snapshot.fonts.map((font) => ({
        ...font,
        paths: [...font.paths],
      })),
      failedResources: this.snapshot.failedResources.map((failure) => ({
        ...failure,
      })),
    };
  }

  subscribe(listener: OfficeRuntimeResourceListener): () => void {
    this.listeners.add(listener);
    return () => this.listeners.delete(listener);
  }

  getVerifiedFontPaths(): string[] {
    return [];
  }

  remainingBytes(): number {
    throw new OfficeRuntimeResourceCompatibilityError(
      "remainingBytes",
      "remainingBytes is unavailable because the streamed CDN has no finite client-side installation plan",
    );
  }

  shouldCheckHealth(now = Date.now()): boolean {
    return now - this.lastHealthCheck > 5 * 60_000;
  }

  prepareForDocumentType(
    _type: OfficeDocumentResourceType,
  ): Promise<OfficeRuntimeResourceProgress> {
    return this.probe("prepare-document");
  }

  prefetchRecommended(): Promise<OfficeRuntimeResourceProgress> {
    return this.probe("prefetch-recommended");
  }

  checkHealth(): Promise<OfficeRuntimeResourceProgress> {
    return this.probe("check-health");
  }

  loadAll(): Promise<OfficeRuntimeResourceProgress> {
    return Promise.reject(new OfficeRuntimeResourceCompatibilityError("loadAll"));
  }

  downloadFontFamily(_id: string): Promise<OfficeRuntimeResourceProgress> {
    return Promise.reject(
      new OfficeRuntimeResourceCompatibilityError("downloadFontFamily"),
    );
  }

  uninstallFontFamily(_id: string): Promise<OfficeRuntimeResourceProgress> {
    return Promise.reject(
      new OfficeRuntimeResourceCompatibilityError("uninstallFontFamily"),
    );
  }

  installFontPreset(
    _preset: OfficeFontPreset,
  ): Promise<OfficeRuntimeResourceProgress> {
    return Promise.reject(
      new OfficeRuntimeResourceCompatibilityError("installFontPreset"),
    );
  }

  pause(): void {
    throw new OfficeRuntimeResourceCompatibilityError("pause");
  }

  resume(): Promise<void> {
    return Promise.reject(new OfficeRuntimeResourceCompatibilityError("resume"));
  }

  cancel(): void {
    throw new OfficeRuntimeResourceCompatibilityError("cancel");
  }

  repair(
    _options: { scope: "required" | "installed" | "all" } = {
      scope: "installed",
    },
  ): Promise<RuntimeCacheProgress> {
    return Promise.reject(new OfficeRuntimeResourceCompatibilityError("repair"));
  }

  plan(_request: ResourcePlanRequest): Promise<ResourcePlan> {
    return Promise.reject(new OfficeRuntimeResourceCompatibilityError("plan"));
  }

  apply(_plan: ResourcePlan): Promise<void> {
    return Promise.reject(new OfficeRuntimeResourceCompatibilityError("apply"));
  }

  checkForUpdates(): Promise<void> {
    return Promise.reject(
      new OfficeRuntimeResourceCompatibilityError("checkForUpdates"),
    );
  }

  maintain(): Promise<RuntimeCacheProgress> {
    return Promise.reject(
      new OfficeRuntimeResourceCompatibilityError("maintain"),
    );
  }

  private async probe(operation: OfficeRuntimeResourceOperation) {
    const urls = [
      getStaticResource().onlyoffice.apiUrl,
      getStaticResource().x2t.script,
    ];
    this.publish({
      ...this.snapshot,
      readiness: "updating",
      operation,
      progress: {
        ...EMPTY_PROGRESS,
        phase: "checking",
        totalFiles: urls.length,
        categories: [],
      },
      error: null,
    });

    try {
      let completedFiles = 0;
      for (const url of urls) {
        const response = await this.fetchImpl(url, {
          cache: "force-cache",
          credentials: "omit",
        });
        if (!response.ok) {
          throw new Error(`Runtime resource probe failed (${response.status}): ${url}`);
        }
        // Consume the response so the browser HTTP cache may retain it.
        await response.arrayBuffer();
        completedFiles += 1;
        this.publish({
          ...this.snapshot,
          progress: {
            ...this.snapshot.progress,
            completedFiles,
          },
        });
      }
      this.lastHealthCheck = Date.now();
      const progress = {
        ...EMPTY_PROGRESS,
        completedFiles: urls.length,
        totalFiles: urls.length,
        categories: [],
      };
      this.publish({
        ...this.snapshot,
        readiness: "ready",
        packs: this.snapshot.packs.map((pack) =>
          pack.id === "core" ? { ...pack, ready: true } : pack,
        ),
        operation: null,
        progress,
        error: null,
      });
      return progress;
    } catch (error) {
      const path = error instanceof Error ? error.message : String(error);
      const progress = {
        ...EMPTY_PROGRESS,
        phase: "error" as const,
        totalFiles: urls.length,
        failedFiles: 1,
        failures: [{ path, reason: "network" }],
        categories: [],
      };
      this.publish({
        ...this.snapshot,
        readiness: "error",
        operation: null,
        progress,
        error: { code: "network", path },
      });
      throw error;
    }
  }

  private publish(snapshot: OfficeRuntimeResourceSnapshot) {
    this.snapshot = snapshot;
    for (const listener of this.listeners) listener(this.getSnapshot());
  }
}

function resolveStaticResourceOptions(
  options: OfficeRuntimeResourceManagerOptions,
): OnlyOfficeStaticResourceOptions | null {
  const cdnOrigin = options.cdnOrigin ?? options.assetBaseUrl;
  if (
    cdnOrigin === undefined &&
    options.onlyofficeVersion === undefined &&
    options.assetManifestDigest === undefined
  ) {
    return null;
  }
  return {
    cdnOrigin: cdnOrigin === null ? null : cdnOrigin?.toString(),
    onlyofficeVersion: options.onlyofficeVersion,
    assetManifestDigest: options.assetManifestDigest,
  };
}

export async function createOfficeRuntimeResourceManager(
  options: OfficeRuntimeResourceManagerOptions = {},
): Promise<OfficeRuntimeResourceManager> {
  return OfficeRuntimeResourceManager.create(options);
}

export {
  getStaticResource,
  registerOnlyOfficeStaticResource,
  resetOnlyOfficeStaticResource,
};
export type { OnlyOfficeStaticResourceOptions };
