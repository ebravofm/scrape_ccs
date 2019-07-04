from utils import scrape_contractors, CONTRATISTAS, PLANILLA, tprint


def main():
    
    ruts_pending = 1
    
    while ruts_pending != 0:
        
        try:

            df = CONTRATISTAS.copy()
            ruts_all = set(df.Rut)

            sets = []
            for sheet in [s for s in PLANILLA.sheet_names if s not in ['Informe RS', 'Lista Contratistas']]:
                sets.append(set(PLANILLA.parse(sheet)['1. Rut'].tolist()))
            ruts_done = set.intersection(*sets)

            ruts_pending = ruts_all - ruts_done

            scrape_contractors(ruts_pending)
            
        except KeyboardInterrupt:
            print()
            tprint('Keyboard Interrupt')
            input()
            sys.exit()
    
    sys.exit()


if __name__ == "__main__":
    main()

