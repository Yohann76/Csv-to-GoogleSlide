# CSV-To-googleslide

### Install & Launch 

```shell
sudo apt-get install python3.7 
git clone https://
cd csv-to-googleslide
python -m pip install -r requirements.txt

python3 csv-to-googleslide.py -h
```

### Need 

- Create JSON credentials_files (https://developers.google.com/workspace/guides/create-credentials?hl=fr#oauth-client-id)
- Activate Google Drive API in console.cloud.google.com
- Activate Google Slide API in console.cloud.google.com
- Create your CSV file for manage your sheets

### Usage & Option

the file parameter (-f) is an obligation in order to specify your csv source

```shell
python3 csv-to-googleslide.py -f "./example/example.csv" (obligatory file option)
python3 csv-to-googleslide.py -f "./example/example.csv" -n "name_of_presentation" (give name to presentation)
```


### CSV Example

|Intitulé formation          |Partie|Intitulé partie  |Objectif d'apprentissage|Intitulé objectif d'apprentissage  |Type objectif d'apprentissage|
|----------------------------|------|-----------------|------------------------|-----------------------------------|-----------------------------|
|Formation pour faire du vélo|1     |Introduction     |1                       |Exemple objectif d'apprentissage 1 |Théorie                      |
|Formation pour faire du vélo|1     |Introduction     |2                       |Exemple objectif d'apprentissage 2 |Théorie                      |
|Formation pour faire du vélo|1     |Introduction     |3                       |Exemple objectif d'apprentissage 3 |Théorie                      |
|Formation pour faire du vélo|2     |Apprendre        |1                       |Exemple objectif d'apprentissage 4 |Théorie                      |
|Formation pour faire du vélo|2     |Apprendre        |2                       |Exemple objectif d'apprentissage 5 |Théorie                      |
|Formation pour faire du vélo|2     |Apprendre        |3                       |Exemple objectif d'apprentissage 6 |Théorie                      |
|Formation pour faire du vélo|2     |Apprendre        |4                       |Exemple objectif d'apprentissage 7 |Théorie                      |
|Formation pour faire du vélo|2     |Apprendre        |5                       |Exemple objectif d'apprentissage 8 |Théorie                      |
|Formation pour faire du vélo|3     |Sans les mains   |1                       |Exemple objectif d'apprentissage 9 |Atelier                      |
|Formation pour faire du vélo|3     |Sans les mains   |2                       |Exemple objectif d'apprentissage 10|Théorie                      |
|Formation pour faire du vélo|4     |Perséverer       |1                       |Exemple objectif d'apprentissage 11|Théorie                      |
|Formation pour faire du vélo|4     |Perséverer       |2                       |Exemple objectif d'apprentissage 12|Théorie                      |
|Formation pour faire du vélo|4     |Perséverer       |3                       |Exemple objectif d'apprentissage 13|Théorie                      |
|Formation pour faire du vélo|5     |Encore           |1                       |Exemple objectif d'apprentissage 14|Théorie                      |
|Formation pour faire du vélo|5     |Encore           |2                       |Exemple objectif d'apprentissage 15|Théorie                      |
|Formation pour faire du vélo|5     |Encore           |3                       |Exemple objectif d'apprentissage 16|Théorie                      |
|Formation pour faire du vélo|5     |Encore           |4                       |Exemple objectif d'apprentissage 17|Théorie                      |
|Formation pour faire du vélo|6     |Plus             |1                       |Exemple objectif d'apprentissage 18|Théorie                      |
|Formation pour faire du vélo|7     |Changer une roue |1                       |Exemple objectif d'apprentissage 19|Atelier                      |
|Formation pour faire du vélo|8     |Nettoyez son vélo|1                       |Exemple objectif d'apprentissage 20|Atelier                      |


(https://www.convertcsv.com/csv-to-markdown.htm)
