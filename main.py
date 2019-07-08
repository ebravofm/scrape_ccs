from utils import scrape_contractors, tprint
import sys
import gpandas as gpd

GSHEET = '1PRYYRoSDazS_aMZ2He_2Okto73QKP4BEl4JwIX2O5dQ'

def main():
    
    ruts_pending = 1
    
    while ruts_pending != 0:
        
        try:
            PLANILLA = gpd.gExcelFile(GSHEET)
            CONTRATISTAS = PLANILLA.parse('Lista Contratistas')

            df = CONTRATISTAS.copy()
            try:
                ruts_all = set(df.Rut)
            except:
                ruts_all = set(df['RUT / Nro.Cliente'])
                
            sets = []
            for sheet in [s for s in PLANILLA.sheet_names if s in ['Curriculum Empresarial', 'Evaluación Comportamiento']]:
                try:
                    sets.append(set(PLANILLA.parse(sheet)['1. Rut'].tolist()))
                except:
                    sets.append(set())
                
            ruts_done = set.intersection(*sets)

            ruts_pending = ruts_all - ruts_done

            scrape_contractors(ruts_pending)
            
        except KeyboardInterrupt:
            print()
            tprint('Keyboard Interrupt')
            input()
            sys.exit()
    
    print()
    tprint('[+] DONE')
    sys.exit()


if __name__ == "__main__":
    main()

