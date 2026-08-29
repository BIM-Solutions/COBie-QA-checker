import type { SpFileService, SpRunHistoryService } from './sharepoint';

/**
 * What the React tree is handed.
 *
 * A plain object rather than a DI container: this web part has two services and
 * one consumer of each, so a registry with tokens and `tryResolve` would be
 * ceremony around a property access. The interface is the seam - a test passes
 * a stub, the workbench could pass a mock - and that is all the indirection the
 * size of this app earns.
 */
export interface CheckerServices {
  readonly files: SpFileService;
  readonly history: SpRunHistoryService;
  /** Library the picker reads COBie files from. Server-relative. */
  readonly sourceFolder: string;
  /** Library exported reports are written to. Server-relative. */
  readonly reportFolder: string;
  /** Signed-in user, recorded against a run. */
  readonly currentUser: string;
  /** False when the site owner has turned run history off in the property pane. */
  readonly historyEnabled: boolean;
}
