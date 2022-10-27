import pandas as pd
import openpyxl
import itertools
from tqdm import tqdm


def dataCleaning(df: pd.DataFrame) -> pd.DataFrame:
    """
    Clean DataFrame
    """
    shoba = df.iloc[13][0]
    mada = df.iloc[15].to_dict()

    df = df.iloc[16:]
    df = df.loc[:, ["Unnamed: 1", "Unnamed: 2", "Unnamed: 3", "Unnamed: 6", "Unnamed: 7",
                    "Unnamed: 10", "Unnamed: 11", "Unnamed: 12", "Unnamed: 16"         
                    ]] 

    df.rename(columns={'Unnamed: 1': mada['Unnamed: 1'],
                   'Unnamed: 2': mada['Unnamed: 2'],
                   'Unnamed: 3': mada['Unnamed: 3'],
                   'Unnamed: 6': mada['Unnamed: 6'],
                   'Unnamed: 7': mada['Unnamed: 7'],
                   'Unnamed: 10': mada['Unnamed: 10'],
                   'Unnamed: 11': mada['Unnamed: 11'],
                   'Unnamed: 12': mada['Unnamed: 12'],
                   'Unnamed: 16': mada['Unnamed: 16']}, inplace=True)

    df = df.reset_index(drop=True)
    names = df["الطالب"].dropna().tolist()
    names = [element.split()[4:] for element in names]
    names = [' '.join(element) for element in names]
    cleaned_names_list = itertools.chain.from_iterable([item, item, item, item, item, item, item] for item in names)
        
    df["الطالب"] = list(cleaned_names_list)
    df["الفصل"] = shoba
    
    return df


def app() -> None:
    '''
    main app running.
    '''
    wb = openpyxl.load_workbook('resources/data.xlsx') 
    sheets_count = len(wb.sheetnames)
    main_df = pd.DataFrame()

    for i in tqdm(range(sheets_count)):
        sheet = pd.read_excel('resources/data.xlsx', sheet_name=i)
        df = dataCleaning(sheet)
        main_df = main_df.append(df)

    main_df.to_csv('resources/combined.csv', encoding='utf-8-sig', index=False)


if __name__ == "__main__":
    app()
