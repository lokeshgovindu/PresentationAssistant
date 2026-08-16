#!/bin/sh
#
# Stamps the commit number into the last component of the extension version, in both places
# that carry it:
#
#   source.extension.vsixmanifest   <Identity ... Version="2026.1.0.9" />
#   source.extension.cs             public const string Version = "2026.1.0.9";
#
# The two must agree: source.extension.cs feeds AssemblyVersion and AssemblyFileVersion, so
# letting them drift means the assembly and the VSIX claim different versions.
#
# The number is the total commit count, not commits since a tag. That makes it monotonic
# whatever happens to tags and releases - and tags do get re-cut. A build number that can
# move backwards is worse than useless for a published extension: the Marketplace will not
# offer the update, and the version is burned for good.
#
# MAJOR.MINOR are left alone; bump those by hand for a release.
#
# Run from the pre-commit hook, where the commit being made is not counted yet, so the
# stamped number is the count plus one. Standalone use:
#
#   .githooks/bump-version.sh            stamp for the next commit
#   .githooks/bump-version.sh --print    print the current version, change nothing
#   .githooks/bump-version.sh --at-head  stamp the count as it stands, for a rebuild
#
set -e

manifest="source.extension.vsixmanifest"
constants="source.extension.cs"

[ -f "$manifest" ] || { echo "bump-version: $manifest not found" >&2; exit 1; }

current=$(sed -n 's/.*<Identity[^>]*Version="\([0-9][0-9.]*\)".*/\1/p' "$manifest" | head -n 1)
[ -n "$current" ] || { echo "bump-version: no version found in $manifest" >&2; exit 1; }

if [ "$1" = "--print" ]; then
    echo "$current"
    exit 0
fi

major=$(echo "$current" | cut -d. -f1)
minor=$(echo "$current" | cut -d. -f2)
build=$(echo "$current" | cut -d. -f3)
[ -n "$minor" ] || minor=0
[ -n "$build" ] || build=0

# A shallow clone reports a count far below the truth, which would stamp a version that goes
# backwards. Refuse rather than do that quietly.
if [ "$(git rev-parse --is-shallow-repository 2>/dev/null)" = "true" ]; then
    echo "bump-version: refusing to stamp from a shallow clone (run: git fetch --unshallow)" >&2
    exit 1
fi

# Anything other than a positive number here would stamp a version lower than one already
# published, which on the Marketplace cannot be undone. Refuse instead of guessing.
count=$(git rev-list --count HEAD 2>/dev/null || true)
case "$count" in
    '' | *[!0-9]*)
        echo "bump-version: cannot determine the commit count - refusing to stamp a version" >&2
        exit 1
        ;;
esac
[ "$count" -gt 0 ] || {
    echo "bump-version: commit count is 0 - refusing to stamp a version" >&2
    exit 1
}

if [ "$1" != "--at-head" ]; then
    count=$((count + 1))          # the commit being made is not counted yet
fi

next="$major.$minor.$build.$count"

sed -i "s/\(<Identity[^>]*Version=\"\)[0-9][0-9.]*\(\"\)/\1$next\2/" "$manifest"

if [ -f "$constants" ]; then
    sed -i "s/\(public const string Version *= *\"\)[0-9][0-9.]*\(\"\)/\1$next\2/" "$constants"
fi

echo "$next"
