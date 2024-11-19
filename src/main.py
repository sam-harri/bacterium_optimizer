import xlwings as xw
import polars as pl
import random
from typing import Dict, Tuple, List


# Function to generate a sample near the midpoint of a range
def generate_midpoint_sample(range_: Tuple[float, float], deviation_percent: float = 0.1) -> float:
    lower, upper = range_
    midpoint = (lower + upper) / 2
    delta = (upper - lower) * deviation_percent
    return random.uniform(midpoint - delta, midpoint + delta)


# Function to run a single simulation
def run_simulation(
    sheet: xw.Sheet,
    params_ranges: Dict[str, Tuple[float, float]],
    params_locations: Dict[str, str],
    validation_params: Dict[str, str],
    results_params: Dict[str, str],
) -> Dict[str, float]:
    # Update input parameters in Excel
    input_values = {}
    for param, cell in params_locations.items():
        value = generate_midpoint_sample(params_ranges[param])
        sheet.range(cell).value = value
        input_values[param] = value

    # Wait for Excel to finish recalculating
    while sheet.api.Application.CalculationState != 0:  # 0 means idle
        pass

    # Check validation cells
    validation_results = {
        param: sheet.range(cell).value for param, cell in validation_params.items()
    }
    if not all(value == "VALID" for value in validation_results.values()):
        raise ValueError(f"Validation failed: {validation_results}")

    # Collect results from output cells
    results = {
        param: sheet.range(cell).value for param, cell in results_params.items()
    }

    # Combine inputs and results into one dictionary
    return {**input_values, **results}


if __name__ == "__main__":
    # Parameters and configurations
    params_locations: Dict[str, str] = {
        "X0": "B12",
        "S0": "B13",
        "Pr0": "B14",
        "Vr": "B15",
    }

    params_ranges: Dict[str, Tuple[float, float]] = {
        "X0": (10, 90),
        "S0": (100, 900),
        "Pr0": (30, 310),
        "Vr": (8000, 72000),
    }

    validation_params: Dict[str, str] = {
        "Pr0_check": "B20",
        "S_check": "B21",
        "YXS_check": "B22",
    }

    results_params: Dict[str, str] = {
        "OPEX_feed_X0": "AB4",
        "OEPX_feed_S0": "AB5",
        "OPEX_feed_Pr0": "AB6",
        "OPEX_utility_cooling": "AB8",
        "OPEX_utility_agitation": "AB9",
        "OPEX_total": "AB10",
        "Revenue_X": "AB14",
        "Revenue_P": "AB15",
        "Revenue_total": "AB16",
        "Profit": "AB18",
    }

    wb_name: str = "data/CHG4381-Group11-ReactorDesignandSimulation.xlsx"
    wb = xw.Book(wb_name)
    sheet = wb.sheets[0]  # Assuming the first sheet is active

    # Initialize Polars DataFrame
    results_df = pl.DataFrame()

    num_simulations = 1000
    save_interval = 100
    save_path = "data/simulation_results.csv"

    try:
        for i in range(1, num_simulations + 1):
            # Run a single simulation
            try:
                simulation_data = run_simulation(
                    sheet,
                    params_ranges,
                    params_locations,
                    validation_params,
                    results_params,
                )
                # Append to Polars DataFrame
                results_df = results_df.vstack(pl.DataFrame([simulation_data]))
            except ValueError as e:
                print(f"Simulation {i} failed: {e}")
                continue

            # Save to file every `save_interval` simulations
            if i % save_interval == 0:
                print(f"Saving results at simulation {i}...")
                results_df.write_csv(save_path)

        # Final save
        results_df.write_csv(save_path)
        print(f"Completed {num_simulations} simulations. Results saved to {save_path}.")
    finally:
        wb.close()
