If (-not (((quser) -replace '^>', '') -replace '\s{2,}', ',' | ConvertFrom-Csv)) {
   shutdown -s -f -t 90 -c "OPGELET: de pc sluit over anderhalve minuut af. Bewaar onmiddellijk uw gegevens!"
}