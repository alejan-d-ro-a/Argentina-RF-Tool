import pandas as pd
import os
from datetime import datetime

class RFDataEngine:
    def __init__(self):
        self.cluster_map = {}
        self.masterbatch_loaded = False

    def _find_column(self, columns, keywords):
        for col in columns:
            col_upper = str(col).strip().upper()
            if all(kw in col_upper for kw in keywords):
                return col
        return None

    def load_masterbatch_to_memory(self, path):
        self.cluster_map.clear()
        try:
            if path.lower().endswith('.csv'):
                df = pd.read_csv(path, dtype=str, encoding='utf-8-sig', on_bad_lines='skip')
            else:
                df = pd.read_excel(path, dtype=str)
            
            col_unico = self._find_column(df.columns, ['UNICO'])
            col_cellname = self._find_column(df.columns, ['CELL', 'NAME'])
            col_cluster = self._find_column(df.columns, ['CLUSTER'])
            
            if col_unico and col_cellname and col_cluster:
                df.fillna("", inplace=True)

                for unico, cluster in zip(df[col_unico], df[col_cluster]):
                    if str(unico).strip():
                        self.cluster_map[str(unico).strip().upper()] = str(cluster).strip()
                        
                for cellname, cluster in zip(df[col_cellname], df[col_cluster]):
                    if str(cellname).strip():
                        self.cluster_map[str(cellname).strip().upper()] = str(cluster).strip()
                
                self.masterbatch_loaded = True
                return True, "Masterbatch loaded. Ready for instant search."
            
            return False, "The selected file does not contain the required columns (UNICO, CELL NAME, CLUSTER)."
        except Exception as e:
            return False, f"Error processing file: {str(e)}"

    def search_cluster_fast(self, query):
        if not self.masterbatch_loaded:
            return "LOAD MASTERBATCH FIRST"
        q = str(query).strip().upper()
        if not q: 
            return ""
        return self.cluster_map.get(q, "NOT FOUND")

    def get_clusters_from_master_ept(self, path):
        try:
            if path.lower().endswith('.csv'):
                df = pd.read_csv(path, usecols=['CLUSTER'], dtype=str, encoding='utf-8-sig')
                return sorted([str(c).strip() for c in df['CLUSTER'].dropna().unique() if str(c).strip()])
            else:
                df = pd.read_excel(path, dtype=str) 
                col_cluster = self._find_column(df.columns, ['CLUSTER'])
                if col_cluster:
                    return sorted([str(c).strip() for c in df[col_cluster].dropna().unique() if str(c).strip()])
                return []
        except Exception:
            return []

    def export_filtered_ept(self, master_ept_path, output_dir, selected_clusters, export_type):
        try:
            if master_ept_path.lower().endswith('.csv'):
                df = pd.read_csv(master_ept_path, dtype=str, encoding='utf-8-sig')
            else:
                df = pd.read_excel(master_ept_path, dtype=str)
            
            col_cluster = self._find_column(df.columns, ['CLUSTER'])
            if not col_cluster:
                return False, "Columna CLUSTER no encontrada en el archivo fuente."

            df['CLUSTER_CLEAN'] = df[col_cluster].astype(str).str.strip()
            df_filtered = df[df['CLUSTER_CLEAN'].isin(selected_clusters)].copy()
            df_filtered.drop(columns=['CLUSTER_CLEAN'], inplace=True)
            
            if df_filtered.empty:
                return False, "No data matched the selected clusters."

            if export_type == "ASSISTANT":
                if 'DlEarfcn' in df_filtered.columns and 'EARFCN' not in df_filtered.columns:
                    df_filtered['EARFCN'] = df_filtered['DlEarfcn']
                elif 'ARFCN' in df_filtered.columns and 'EARFCN' not in df_filtered.columns:
                    df_filtered['EARFCN'] = df_filtered['ARFCN']
                    
                if 'CLUSTER' in df_filtered.columns and 'Cluster' not in df_filtered.columns:
                    df_filtered['Cluster'] = df_filtered['CLUSTER']

                cols = ['Longitude', 'Latitude', 'Azimuth', 'eNodeB Name', 'eNodeB ID', 'Cell Name', 'Cell ID', 'EARFCN', 'PCI', 'isOutdoor', 'Region', 'Cluster', 'Mesh', 'Pilot Phase Network', 'MCC', 'MNC', 'Height', 'Mechanical Downtilt', 'Electrical Downtilt', 'Sectorization', 'TAC']
                for c in cols:
                    if c not in df_filtered.columns:
                        df_filtered[c] = ""
                
                col_cellname = self._find_column(df_filtered.columns, ['CELL', 'NAME'])
                if col_cellname:
                    df_filtered['Sectorization'] = df_filtered[col_cellname].astype(str).str.strip().str[-1]
                df_filtered = df_filtered[cols]

            elif export_type == "ACP":
                cols = ['TAC', 'Physical Site Name', 'eNodeB ID', 'Cell ID', 'eNodeB Name', 'Cell Name', 'Sector Split Group Id', 'Active', 'Longitude', 'Latitude', 'Azimuth', 'Beam Azimuth', 'Beam Azimuth Offset', 'Height', 'Mechanical Downtilt', 'Electrical Downtilt', 'Antenna Model', 'Antenna Pattern', 'PCI', 'DlEarfcn', 'DlBandwidth', 'RS Power', 'PA', 'PB', 'Number of Transmission Antenna Ports', 'Number of Transmission Antennas', 'PDSCH Actual Load(DL)', 'isOutdoor', 'MCC', 'MNC', 'Local Cell ID', 'Max Power', 'Main Resolution', 'Scenario', 'Site Type', 'Sectorization', 'Status', 'RRU Serial Number']
                for c in cols:
                    if c not in df_filtered.columns:
                        df_filtered[c] = ""
                
                col_cellname = self._find_column(df_filtered.columns, ['CELL', 'NAME'])
                if col_cellname:
                    df_filtered['Sectorization'] = df_filtered[col_cellname].astype(str).str.strip().str[-1]
                df_filtered = df_filtered[cols]

            cols_to_check = [col for col in df_filtered.columns if col != 'MNC']
            df_filtered.drop_duplicates(subset=cols_to_check, inplace=True)
            
            safe_names = [str(c).replace(" ", "_") for c in selected_clusters]
            cluster_str = "_".join(safe_names[:4]) + "_AND_MORE" if len(safe_names) > 4 else "_".join(safe_names)
                
            stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"{export_type}_EPT_{cluster_str}_{stamp}.xlsx"
            out_path = os.path.join(output_dir, filename)
            
            df_filtered.to_excel(out_path, index=False)
            return True, out_path
        except PermissionError:
            return False, "Permission denied. Asegúrate de que el archivo no esté abierto."
        except Exception as e:
            return False, str(e)
