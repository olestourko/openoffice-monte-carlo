#### Start Open Office / Libre Office
 `soffice <calc_file.ods> --accept=socket,host=localhost,port=2002;urp`
 
#### Run the Monte Carlo script
Make sure to use Python 3.x!

`python3 monte_carlo.py --model_sheet 'model' --output_cell 'A17' --variables_sheet 'risk_variables'`