"use client";

import { useSqlGenerator } from "./use-sql-generator";
import { GeneratorControls } from "./generator-controls";
import { GeneratorFilesTable } from "./generator-files-table";
import { GeneratorHero } from "./generator-hero";
import { GeneratorParametersTable } from "./generator-parameters-table";
import { GeneratorStatus } from "./generator-status";

export function SqlGeneratorPage() {
  const generator = useSqlGenerator();

  return (
    <main className="relative min-h-screen overflow-hidden bg-[radial-gradient(circle_at_top_left,rgba(218,255,239,0.85),transparent_30%),radial-gradient(circle_at_top_right,rgba(255,231,204,0.8),transparent_28%),linear-gradient(160deg,#f7f5ef,#e8efe6)] text-slate-900">
      <div
        aria-hidden="true"
        className="pointer-events-none absolute inset-0 -z-10 bg-[linear-gradient(rgba(19,52,60,0.045)_1px,transparent_1px),linear-gradient(90deg,rgba(19,52,60,0.045)_1px,transparent_1px)] bg-size-[28px_28px] mask-[radial-gradient(circle_at_center,#000_38%,transparent_85%)]"
      />

      <div className="mx-auto flex w-full max-w-7xl flex-col gap-5 px-4 py-5 sm:px-6 lg:px-8">
        <GeneratorHero
          isProcessing={generator.isProcessing}
          canExport={generator.canExport}
          onProcess={generator.processFile}
          onDownloadZip={generator.downloadZip}
          onDownloadAll={generator.downloadAllUncompressed}
          onOpenAll={generator.openAllInBrowser}
        />

        <GeneratorControls
          file={generator.file}
          secEjecutable={generator.secEjecutable}
          roleName={generator.roleName}
          includeMenuPredecessors={generator.includeMenuPredecessors}
          sortBy={generator.sortBy}
          onFileChange={generator.setFile}
          onSecEjecutableChange={generator.setSecEjecutable}
          onRoleNameChange={generator.setRoleName}
          onIncludeMenuPredecessorsChange={generator.setIncludeMenuPredecessors}
          onSortChange={generator.handleSortChange}
        />

        <GeneratorStatus
          workbookName={generator.workbookName}
          status={generator.status}
          generatedContext={generator.generatedContext}
          generatedPairs={generator.generatedPairs}
          generatedFiles={generator.generatedFiles}
          otherFiles={generator.otherFiles}
        />

        <GeneratorFilesTable files={generator.otherFiles} onDownload={generator.downloadFile} onOpen={generator.openFile} />

        <GeneratorParametersTable pairs={generator.generatedPairs} onDownloadPair={generator.downloadPair} onOpenPair={generator.openPairInBrowser} />
      </div>
    </main>
  );
}
