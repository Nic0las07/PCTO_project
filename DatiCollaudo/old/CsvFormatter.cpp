#include <iostream>
#include <filesystem>
#include <fstream>
#include <vector>
#include <map>

bool moveFile(std::string filename, std::string destination);

int main() {

    //Controllo di tutti i file e cartelle presenti

    for (const auto& entry : std::filesystem::directory_iterator(".")) {
        if (entry.is_regular_file()) {
            std::string filename{ entry.path().filename().string() };
            int num_dots{0};

            //Controllo correttezza nomi file ed eventuale modifica

            for (int index{0}; index < filename.length(); index++) {
                if (filename[index] == '.') {
                    num_dots++;
                }

                if (!std::isalnum(filename[index]) && filename[index] != '_') {
                    if (filename[index] == '.' && num_dots <= 1) continue;

                    filename[index] = '_';
                }
            }

            //Analisi di tutti i file .csv presenti

            if (filename.find(".csv") != -1) {
                std::cout << "\n\n\n" << "Avvio analisi del file: " << filename << "\n\n";
                
                std::ifstream input_file(filename); 

                if (!input_file) {
                    std::cout << "Errore nell'apertura del file" << std::endl;
                    continue;
                } else {
                    std::cout << "File trovato ed aperto con successo" << std::endl;
                }

                std::string line{""};
                std::vector<std::string> line_parts;
                std::vector<std::string> data_type;
                std::vector<std::string> data;

                char separator{','};
                int row_num{0};

                //Lettura ed analisi di ogni riga del file corrente

                while (std::getline(input_file, line) && row_num != 2) {
                    if(line == "") continue;

                    int start_index{0};
                    line_parts.clear();

                    //Separazione dei diversi campi dati di una riga

                    for (int index{0}; index <= line.length(); index++) {
                        if (index == line.length() || line[index] == separator) {
                            line_parts.push_back(line.substr(start_index, index - start_index));
                            start_index = index + 1;
                        }
                    }

                    //Inserimento dei dati e dei loro campi identificativi all'interno di una struttura

                    if (line_parts.at(0) == "TIME") {
                        for (std::string str : line_parts) {
                            data_type.push_back(str);
                        }
                    } else {
                        for (std::string str : line_parts) {
                            data.push_back(str);
                        }
                    }

                    row_num++;
                }

                input_file.close();

                //Controlli per la validità del file

                if (row_num != 2) {
                    std::cout << "Non ci sono abbastanza righe" << std::endl;
                    continue;
                } else {
                    std::cout << "File analizzato con successo e dati memorizzati" << std::endl;
                }

                if (data_type.size() != data.size()) {
                    std::cout << "Il numero di campi relativi agli identificativi dei dati non coincide con quello dei campi contenenti i dati";
                    continue;
                }

                //Conversione dei dati che ne hanno bisogno

                for (int data_field{0}; data_field < data.size(); data_field++) {
                    std::string new_data{ data.at(data_field) };
                    while (new_data.find('.') != -1) {
                        new_data.erase(new_data.find('.'), 1);
                    }
                    data.at(data_field) = new_data;
                }

                std::string fields_to_modify[] = {"Temp", "Pressione", "Vel", "Portata"};

                for (int data_field{0}; data_field < data.size(); data_field++) {

                    for(auto field : fields_to_modify){
                        if(data_type.at(data_field).find(field) != -1){
                            double value = std::stod(data.at(data_field)) * 0.000001;
                            std::string new_data = std::to_string(value);
                            
                            if(new_data.find('.') != -1){
                                new_data = new_data.substr(0, new_data.find('.') + 2);
                            }

                            data.at(data_field) = new_data;
                            
                        }
                    }

                }

                std::cout << "Dati convertiti con successo" << std::endl;

                //Scrittura di tutti i dati ottenuti all'interno del file .csv

                std::ofstream output_file(filename);

                std::map<std::string, std::string> measurement_units;

                measurement_units["Temp"] = "'C";
                measurement_units["Pressione"] = "bar";
                measurement_units["Pos"] = "stps";
                measurement_units["Vel"] = "rps";
                measurement_units["Portata"] = "l/h";

                for (int data_field{0}; data_field < data_type.size(); data_field++) {
                    line.clear();

                    line += data_type.at(data_field) + ";";
                    line += data.at(data_field) + ";";

                    //Assegnazione delle unità di misura per alcuni dati
                    
                    for(auto mu : measurement_units){
                        if(data_type.at(data_field).find(mu.first) != -1){
                            line += mu.second + ";";
                        }
                    }

                    if (data_field == data_type.size() - 1) {
                        output_file << line;
                    } else {
                        output_file << line << std::endl;
                    }
                }

                std::cout << "Inserimento nel file dei dati convertiti avvenuto con successo" << std::endl;
                std::cout << "Riformattazione avvenuta con successo" << std::endl;

                output_file.close();

                //Spostamento del file csv finale in una cartella apposita

                if(moveFile(filename, "formatted_csv")){
                    std::cout << "File spostato con successo" << std::endl;
                }else{
                    std::cout << "Spostamento file non riuscito" << std::endl;
                }
            }
        }
    }

    return 0;
}

bool moveFile(std::string filename, std::string destination){

    try {
        std::filesystem::create_directory(destination);

        std::filesystem::rename(filename, destination + "/" + filename);
    } catch (std::filesystem::filesystem_error e) {
        return false;
    }
    return true;
}