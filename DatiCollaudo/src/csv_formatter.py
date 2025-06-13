from pathlib import Path
import csv
import os

directory = Path('raw_csv_files')

for current_file in directory.iterdir():

    if current_file.suffix == '.csv':

        print('\n\n' + current_file.name + '\n')

        fields_name = []
        fields_data = []
        csv_data = {}

        with open(f'{directory.name}\\{current_file.name}', newline = '', encoding = 'utf-8') as csvfile:
            reader = csv.reader(csvfile, delimiter = ',')

            for row in reader:

                if len(row) != 0:

                    if row[0].lower() == 'time': 
                        fields_name = row
                    elif len(fields_data) == 0:
                        fields_data = row

        if len(fields_data) != len(fields_name):
            print('Il numero di dati non corrisponde con il numero di campi')
            continue 
        elif len(fields_name) == 0:
            print('Il file è vuoto')
            continue
        else: print('Analisi file e memorizzazione dati effettuata con successo')
        
        for index in range(0, len(fields_name)): csv_data[fields_name[index]] = fields_data[index]
        
        fields_to_modify = ['temp', 'pressione', 'vel', 'portata']

        fields_modified = False

        for key, value in csv_data.items():

            for field_to_modify in fields_to_modify:

                if field_to_modify in key.lower():

                    csv_data[key] = str(float(value))
                    fields_modified = True
                    break

        if fields_modified == True:
            print('Conversione dati effettuata con successo')
        else:
            print('Non ci sono stati dati da convertire')
        
        

        measurement_units = {}
        measurement_units['temp'] = '\'C'
        measurement_units['pressione'] = 'bar'
        measurement_units['pos'] = 'stps'
        measurement_units['vel'] = 'rps'
        measurement_units['portata'] = 'l/h'

        new_file_name = current_file.name.replace(current_file.suffix, '') + '_formatted.csv'
        destination_directory = 'formatted_csv'
        os.makedirs(destination_directory, exist_ok=True)

        with open(os.path.join(destination_directory, new_file_name), 'w', newline = '', encoding = 'utf-8') as csvfile:
            writer = csv.writer(csvfile, delimiter = ';')

            for field_name, field_data in csv_data.items():
                measurement_unit = ''

                for type_unit, unit in measurement_units.items():

                    if type_unit in field_name.lower():
                        measurement_unit = unit

                field_name = field_name.replace('_Mevo', '')    

                writer.writerow([field_name, field_data, measurement_unit])

        print('Creazione del nuovo file csv formattato e contenente i dati avvenuta con successo')