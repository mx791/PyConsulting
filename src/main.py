from dotenv import load_dotenv
from financial_presentation import make_financial_pres

if __name__ == "__main__":
    load_dotenv()
    symbols = ["DTE.DE", "SU.PA", "SAP", "AI.PA", "AIR.PA", "HO.PA", "OR.PA", "TTE"]
    make_financial_pres(symbols, "Charles Lrz", "./out/out.pptx")
