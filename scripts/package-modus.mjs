import { execFileSync } from "node:child_process";
import { copyFileSync, cpSync, mkdirSync, readFileSync, rmSync, writeFileSync } from "node:fs";

const root = new URL("../", import.meta.url);
const output = new URL("dist/modus-package/", root);
const git = (...args) => execFileSync("git", args, { cwd: root, encoding: "utf8" });
if (git("status", "--porcelain").trim()) throw new Error("Commit the source before packaging a release.");
const sourceRevision = git("rev-parse", "HEAD").trim();
const upstreamRevision = "873cd65bf3d04c4d1ff4f61263c7a2544b6571d0";
const metadata = JSON.parse(readFileSync(new URL("packages/react-xlsx/package.json", root), "utf8"));
metadata.repository.url = "git+https://github.com/modus-audit/react-xlsx.git";
metadata.modusSource = sourceRevision;
metadata.modusUpstream = upstreamRevision;
delete metadata.scripts;
delete metadata.devDependencies;
rmSync(output, { recursive: true, force: true });
mkdirSync(output, { recursive: true });
cpSync(new URL("packages/react-xlsx/dist/", root), new URL("dist/", output), { recursive: true });
copyFileSync(new URL("LICENSE", root), new URL("LICENSE", output));
writeFileSync(new URL("package.json", output), `${JSON.stringify(metadata, null, 2)}\n`);
writeFileSync(new URL("SOURCE.patch", output), git("diff", upstreamRevision, sourceRevision, "--", ".", ":(exclude).github/workflows"));
writeFileSync(new URL("README.md", output), `# Modus React XLSX distribution

Version ${metadata.version}, based on upstream 0.16.4.

Source: https://github.com/modus-audit/react-xlsx/commit/${sourceRevision}

The source repository's MODUS_PATCHES.md documents the patches, validation, and build commands. SOURCE.patch contains the complete code changes from upstream; the fork's existing workflows are retained.
`);
