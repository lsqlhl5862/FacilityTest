echo Run Tesseract for Training.. 
tesseract.exe num.font.exp0.tif num.font.exp0  -l eng --psm 7 nobatch box.train 
 
echo Compute the Character Set.. 
unicharset_extractor.exe num.font.exp0.box 
shapeclustering -F font_properties -U unicharset -O num.unicharset num.font.exp0.tr
mftraining -F font_properties -U unicharset -O num.unicharset num.font.exp0.tr 


echo Clustering.. 
cntraining.exe num.font.exp0.tr 

echo Rename Files.. 
rename normproto num.normproto 
rename inttemp num.inttemp 
rename pffmtable num.pffmtable 
rename shapetable num.shapetable  
rename unicharset  normal.unicharset

echo Create Tessdata.. 
combine_tessdata.exe num. 

echo. & pause