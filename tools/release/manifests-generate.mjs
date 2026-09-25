#!/usr/bin/env node
import {
  alphaManifest,
  alphaProfile,
  cimbProdProfile,
  developmentManifest,
  ensureDir,
  generatedManifestPath,
  generatedOneNoteManifestPath,
  generatedOfficeXmlManifestPath,
  multiHostOfficeXmlManifest,
  oneNoteManifest,
  outlookXmlManifest,
  profileFromArgs,
  parseArgs,
  releaseConfig,
  repoRoot,
  taskPaneXmlManifest,
  writeJson,
} from './common.mjs';
import { dirname, join } from 'node:path';
import { writeFileSync } from 'node:fs';

const args = parseArgs();
const profile = profileFromArgs(args);

try {
  const cfg = releaseConfig(profile);
  const manifest = profile === alphaProfile ? alphaManifest(cfg) : developmentManifest(cfg);
  const out = generatedManifestPath(profile);
  ensureDir(dirname(out));
  writeJson(out, manifest);
  console.log(`generated ${out}`);
  const write = (path, text) => {
    writeFileSync(path, text);
    console.log(`generated ${path}`);
  };
  ensureDir(join(repoRoot, 'dist', 'manifests'));
  if (profile === 'development') {
    write(generatedOneNoteManifestPath(profile), oneNoteManifest(cfg));
    for (const surface of ['word', 'excel', 'powerpoint']) {
      write(generatedOfficeXmlManifestPath(profile, surface), taskPaneXmlManifest(cfg, surface));
    }
    write(generatedOfficeXmlManifestPath(profile, 'office'), multiHostOfficeXmlManifest(cfg));
    write(generatedOfficeXmlManifestPath(profile, 'outlook'), outlookXmlManifest(cfg));
  } else if (profile === cimbProdProfile) {
    // Production ships the unified package plus centralized-deployment XML for the chosen apps.
    const has = (s) => cfg.surfaces.includes(s);
    if (['word', 'excel', 'powerpoint'].some(has)) {
      write(generatedOfficeXmlManifestPath(profile, 'office'), multiHostOfficeXmlManifest(cfg));
    }
    if (has('outlook')) {
      write(generatedOfficeXmlManifestPath(profile, 'outlook'), outlookXmlManifest(cfg));
    }
    if (has('onenote')) write(generatedOneNoteManifestPath(profile), oneNoteManifest(cfg));
  }
} catch (err) {
  if (err?.code === 'BLOCKED_EXTERNAL') {
    console.error(`BLOCKED_EXTERNAL: ${err.message}`);
    process.exit(2);
  }
  console.error(err instanceof Error ? err.message : String(err));
  process.exit(1);
}
