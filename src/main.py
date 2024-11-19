import time
import xlwings as xw # type: ignore[import-untyped]
import polars as pl
import random
from typing import Dict, Tuple
from loguru import logger


def sampling_function(range_: Tuple[float, float]) -> float:
    upper, lower = range_
    return random.uniform(upper, lower)


def run_simulation(
    sheet: xw.Sheet,
    params_stats: Dict[str, Tuple[float, float]],
    params_locations: Dict[str, str],
    validation_params: Dict[str, str],
    results_params: Dict[str, str],
) -> Dict[str, float | str]:
    input_values = {}
    for param, cell in params_locations.items():
        value = sampling_function(params_stats[param])
        sheet[cell].value = value
        input_values[param] = value

    IS_IDLE = 0
    while sheet.api.Application.CalculationState != IS_IDLE:
        pass

    validation_results = {
        param: sheet.range(cell).value for param, cell in validation_params.items()
    }
    if not all(value == "VALID" for value in validation_results.values()):
        raise ValueError(f"Validation failed: {validation_results}")

    results : Dict[str, float | str]= {param: sheet.range(cell).value for param, cell in results_params.items()}

    return {**input_values, **validation_results, **results}


if __name__ == "__main__":
    params_locations: Dict[str, str] = {
        "X0": "B12",
        "S0": "B13",
        "Pr0": "B14",
        "Vr": "B15",
    }

    params_stats: Dict[str, Tuple[float, float]] = {
        "X0": (45, 55),
        "S0": (460, 490),
        "Pr0": (150, 200),
        "Vr": (35_000, 45_000),
    }

    validation_params: Dict[str, str] = {
        "Pr0_check": "B20",
        "S_check": "B21",
        "YXS_check": "B22",
    }

    results_params: Dict[str, str] = {
        "operation_total_time": "B25",
        "operation_avg_batch": "B26",
        "operation_batch_time": "B27",
        "operation_num_batches": "B28",
        "operation_stationary_start": "B29",
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
    wb : xw.Book = xw.Book(wb_name)
    sheet : xw.Sheet = wb.sheets["Reactor"]
    time.sleep(1)

    results_df : pl.DataFrame = pl.DataFrame()

    num_simulations : int = 10_000
    save_interval : int = 100
    save_path :str = "data/simulation_results.csv"

    try:
        for i in range(1, num_simulations + 1):
            try:
                now = time.time()
                simulation_data = run_simulation(
                    sheet,
                    params_stats,
                    params_locations,
                    validation_params,
                    results_params,
                )
                logger.info(
                    f"SUCCESFULL  | Simulation {i} completed in {1000*(time.time() - now):.2f} miliseconds."
                )
                results_df = results_df.vstack(pl.DataFrame([simulation_data]))
            except ValueError:
                logger.info(
                    f"UNSUCCESFUL | Simulation {i} completed in {1000*(time.time() - now):.2f} miliseconds."
                )
            finally:
                time.sleep(0.1)
                if i % save_interval == 0:
                    logger.info(f"SAVING | Saving simulation result checkpoint {i}...")
                    results_df.write_csv(save_path)

        results_df.write_csv(save_path)
        logger.info(f"COMPLETED | {num_simulations} simulations. Results saved to {save_path}.")
    finally:
        wb.close()
