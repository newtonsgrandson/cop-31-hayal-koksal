#!/bin/bash
# Renders the 6 scenes frame by frame and assembles the looping Full HD video.
set -euo pipefail
W="$DS_WORK"; ROOT="$(dirname "$W")"; SCR="$(dirname "$0")"
FPS=25; DUR=13; XF=0.5
mkdir -p "$W/seg"
for n in ${SEGS:-1 2 3 4 5 6}; do
  rm -rf "$W/frames"; node "$SCR/ds_shots.js" frames "$W/html/slide$n.html" "$W/frames" $DUR $FPS
  if [ $n -eq 1 ] || [ $n -eq 6 ]; then
    ffmpeg -loglevel error -y -i "$ROOT/tpl/ppt/media/media1.mp4" -framerate $FPS -i "$W/frames/%05d.png" \
      -filter_complex "[0:v]fps=$FPS,scale=1920:1080,tpad=stop_mode=clone:stop_duration=$DUR,trim=duration=$DUR,setpts=PTS-STARTPTS[bg];[bg][1:v]overlay=format=auto,format=yuv420p[v]" \
      -map "[v]" -t $DUR -c:v libx264 -crf 14 -preset medium "$W/seg/s$n.mp4"
  else
    ffmpeg -loglevel error -y -stream_loop -1 -i "$ROOT/tpl/ppt/media/media2.mp4" -framerate $FPS -i "$W/frames/%05d.png" \
      -filter_complex "[0:v]fps=$FPS,scale=1920:1080,trim=duration=$DUR,setpts=PTS-STARTPTS[bg];[bg][1:v]overlay=format=auto,format=yuv420p[v]" \
      -map "[v]" -t $DUR -c:v libx264 -crf 14 -preset medium "$W/seg/s$n.mp4"
  fi
  echo "segment $n done"
done
rm -rf "$W/frames"
O1=$(echo "$DUR-$XF" | bc); O2=$(echo "2*($DUR-$XF)" | bc); O3=$(echo "3*($DUR-$XF)" | bc); O4=$(echo "4*($DUR-$XF)" | bc); O5=$(echo "5*($DUR-$XF)" | bc)
TOTAL=$(echo "6*$DUR-5*$XF" | bc); FO=$(echo "$TOTAL-0.5" | bc)
ffmpeg -loglevel error -y -i "$W/seg/s1.mp4" -i "$W/seg/s2.mp4" -i "$W/seg/s3.mp4" -i "$W/seg/s4.mp4" -i "$W/seg/s5.mp4" -i "$W/seg/s6.mp4" \
  -f lavfi -i anullsrc=r=48000:cl=stereo \
  -filter_complex "[0:v][1:v]xfade=transition=fade:duration=$XF:offset=$O1[a];[a][2:v]xfade=transition=fade:duration=$XF:offset=$O2[b];[b][3:v]xfade=transition=fade:duration=$XF:offset=$O3[c];[c][4:v]xfade=transition=fade:duration=$XF:offset=$O4[d];[d][5:v]xfade=transition=fade:duration=$XF:offset=$O5,fade=t=out:st=$FO:d=0.5:color=white,format=yuv420p[v]" \
  -map "[v]" -map 6:a -t $TOTAL -r $FPS -c:v libx264 -profile:v high -crf 18 -preset slow -pix_fmt yuv420p -c:a aac -b:a 96k -movflags +faststart \
  "$W/COP31_Climate_IMECE_Digital_Screen_EN_1920x1080.mp4"
echo "VIDEO DONE total=$TOTAL"
