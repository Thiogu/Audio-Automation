
#!/usr/bin/env bash
#set -euo pipefail

VOICE="${VOICE:-Matthew}"       # change after you list voices
FORMAT="${FORMAT:-mp3}"        # mp3 | ogg_vorbis | pcm
ENGINE="${ENGINE:-standard}"   # use 'neural' if you select an NTTS-only voice
INPUT_JSON="${INPUT_JSON:-/home/cloudshell-user/test.json}"
OUTDIR="${OUTDIR:-/home/cloudshell-user/testFolder}" 

for cmd in jq aws; do
  command -v "$cmd" >/dev/null 2>&1 || { echo "❌ Missing: $cmd"; exit 1; }
done

[[ -f "$INPUT_JSON" ]] || { echo "❌ Not found: $INPUT_JSON"; exit 1; }
mkdir -p "$OUTDIR"

COUNT=$(jq '. | length' "$INPUT_JSON")
echo "🔄 Slides to process: $COUNT (VOICE=$VOICE, ENGINE=$ENGINE)"

jq -c '.[]' "$INPUT_JSON" | while read -r obj; do
  idx=$(echo "$obj" | jq -r '.index')
  text=$(echo "$obj" | jq -r '.text')
  [[ -z "$text" || "$text" == "null" ]] && { echo "⚠️  Slide $idx has empty notes; skipping"; continue; }

  outfile="$OUTDIR/slide_${idx}.${FORMAT}"
  aws polly synthesize-speech \
    --output-format "$FORMAT" \
    --voice-id "$VOICE" \
    --text "$text" \
    --engine "$ENGINE" \
    "$outfile"
  [[ -s "$outfile" ]] && echo "✔️  $outfile" || echo "❌ Empty output for slide $idx"
done

echo "✅ Done. Files are in: $OUTDIR"

