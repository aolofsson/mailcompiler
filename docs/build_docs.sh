#!/usr/bin/env bash
#
# build_docs.sh - regenerate all derived documentation artifacts from source.
#
# Lives in docs/; renders every Graphviz .dot alongside it to .svg and .png.
# Add new source -> artifact conversions here as the docs grow.

set -euo pipefail

cd "$(dirname "$0")"

command -v dot >/dev/null 2>&1 || {
    echo "error: graphviz 'dot' not found; install graphviz" >&2
    exit 1
}

shopt -s nullglob
dots=(*.dot)
if [ ${#dots[@]} -eq 0 ]; then
    echo "no .dot sources in docs/"
else
    for src in "${dots[@]}"; do
        base="${src%.dot}"
        echo "dot: $src -> $base.svg, $base.png"
        dot -Tsvg "$src" -o "$base.svg"
        dot -Tpng "$src" -o "$base.png"
    done
fi

echo "docs built."
