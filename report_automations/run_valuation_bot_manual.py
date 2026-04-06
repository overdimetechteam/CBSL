"""
Manual Initial Valuation Bot Runner
Run this script to fetch initial valuations for specific contracts manually.

Usage:
    python run_valuation_bot_manual.py HQDL250007660 HQDL250007700 ...
"""

import sys
import os

# Add bots directory to path
sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(__file__)), 'bots'))

from IA_Working_Initial_valuation_bot import IAWorkingInitialValuationBot

def main():
    if len(sys.argv) < 2:
        print("Usage: python run_valuation_bot_manual.py <contract_no1> <contract_no2> ...")
        print("Example: python run_valuation_bot_manual.py HQDL250007660 HQDL250007700")
        sys.exit(1)

    contracts = sys.argv[1:]
    print(f"Running Initial Valuation Bot for {len(contracts)} contracts...")
    print(f"Contracts: {contracts}")

    try:
        # Initialize bot
        bot = IAWorkingInitialValuationBot()

        # Run bot
        results = bot.run_valuation_bot(contracts)

        # Display results
        print("\n" + "="*80)
        print("RESULTS")
        print("="*80)

        if isinstance(results, dict):
            for contract, value in results.items():
                print(f"{contract}: {value}")
        else:
            print(f"Unexpected result type: {type(results)}")
            print(results)

        print("="*80)

        return results

    except Exception as e:
        print(f"\nERROR: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

if __name__ == "__main__":
    main()