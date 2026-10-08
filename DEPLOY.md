# Deployment

How a release of the `ppt` and `ppt-prepare` skills is built and delivered.
The pptc CLI itself ships bundled inside the `ppt` skill.

## Channels

| Channel | Artifact | Templates | Where it goes |
|---|---|---|---|
| GitHub release (public) | `ppt.zip`, `ppt-prepare.zip`, `ppt-skills-<version>.zip` | neutral only | GitHub release assets |
| Claude app (company) | `ppt-msg.zip`, `ppt-prepare-msg.zip` | msg templates | uploaded by hand |
| Company marketplace | `ppt-skills-msg-<version>.zip` | msg templates | marketplace PR |

All ZIPs are built by `npm run skill:zip` into `deploy/` (gitignored).

> The msg templates must never be published. They live in `private-templates/`
> (gitignored) and only ever land in the `*-msg*` ZIPs. A GitHub release carries
> the neutral template only; never attach a `*-msg*` ZIP to it.

## Versions

- `package.json` is the single source of the **skill version**.
- `scripts/sync-versions.mjs` (run by `plugin:sync`) writes it into
  `plugin/skills/ppt/VERSION` and `plugin/skills/ppt-prepare/VERSION` only.
- The **plugin version** (`plugin.json` of the company plugin `common`) belongs to
  the marketplace owner and is bumped by hand; it is not derived from the skills.
  `plugin.json` and `marketplace.json` in this repo carry no version.

## ZIP layouts

- **Plugin layout** (`ppt-skills*.zip`): `skills/ppt`, `skills/ppt-prepare`,
  `meta/control.md`. The skills import `../../meta/control.md`, so keep the
  folders together.
- **Stand-alone** (`ppt*.zip`, `ppt-prepare*.zip`): the skill folder at the ZIP
  root. There is no `../../meta`, so the `@...control.md` import line in `SKILL.md`
  is replaced by the inlined content of `meta/control.md`.
- **msg variant**: `private-templates/*.potx|pptx|md` replace the neutral template
  (`neutral-template.pptx` and its sidecar are dropped), and a `metadata:` block
  (business/technical owner and the skill version) is stamped into each `SKILL.md`
  front matter, as the company marketplace requires. The owners come from
  `private-templates/skill-metadata.json`:

  ```json
  {
      "business-owner": "<names>",
      "technical-owner": "<names>"
  }
  ```

  The msg ZIPs are built only if `private-templates/` exists (or `--from <dir>`).

## Release checklist

1. Bump the version (`node scripts/version-bump.mjs patch`), or edit
   `package.json`; write the `CHANGELOG.md` section.
2. `npm test` and `npm run lint`; both must be green.
3. `npm run skill:zip` (runs `plugin:sync`: bundle, `VERSION` files; then all ZIPs).
4. Commit, tag (`git tag -a <version>`), `git push origin main <version>`.
5. GitHub release with the **neutral** ZIPs only:
   `gh release create <version> --verify-tag --title "<version>" --notes-from-tag deploy/ppt.zip deploy/ppt-prepare.zip deploy/ppt-skills-<version>.zip`
6. Hand the msg ZIPs to the company channels (see below).

## Company marketplace

The company marketplace repo is cloned to `deploy/msg-marketplace/` (inside the
gitignored `deploy/`). It uses a different GitHub account than this repo, see
**Accounts**.

Per release, in a branch of that repo:

1. Unpack `ppt-skills-msg-<version>.zip` into `plugin/common/` (`skills/`, `meta/`).
   The `SKILL.md` files then match the ones this repo generates.
2. Align the CLI source `src/common/pptc/` with this repo: `package.json` version and
   dependencies, `package-lock.json`, `CHANGELOG.md`. `src/` itself usually does not
   change between patch releases. The repo's tests there use the company paths; keep
   them.
3. Bump the **plugin** version in `plugin/common/.claude-plugin/plugin.json` by hand
   and add the German per-skill `CHANGELOG.md` entries (`plugin/common/skills/*/`).
4. Run `node scripts/build.mjs` there (after `npm ci` in `src/common/pptc`). It builds
   `pptc.mjs` and stamps `VERSION` and `SKILL.md` `metadata.version` from the
   `src/common/pptc` version. It must not change anything else; if `git status` shows
   a diff in `plugin/`, the source is out of sync with the committed bundle.
5. Open the PR. CI (`build.yml`) rebuilds on every change under `src/common/pptc/`, so
   the source there must be at the released version, otherwise CI overwrites the
   bundle and the `VERSION` files with the older source.

## Accounts

Two GitHub accounts are logged in with `gh` (`gh auth status`, `gh auth switch`):
the personal one for this repo and the company one for the marketplace.

- `git push` in this repo uses the `gh` login (repo-local
  `credential.helper = !gh auth git-credential`), so switch to the personal account
  first.
- The marketplace clone uses the macOS keychain (company account).
- The company account is an SSO account; browser login happens via the company SSO
  page, then `gh auth login` with the device code in that browser session.
