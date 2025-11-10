import os
import time
import itertools
import pythoncom
import win32com.client
import pandas as pd
import numpy as np
from multiprocessing import Pool, cpu_count
from tqdm import tqdm
from SALib.sample import saltelli, morris, latin, fast_sampler
from SALib.analyze import sobol, morris as morris_analyze, delta, fast
from scipy.stats.qmc import Sobol as QmcSobol, Halton
import matplotlib.pyplot as plt
import seaborn as sns
import threading
import tkinter as tk
from tkinter import messagebox
import subprocess
import psutil
from scipy.optimize import minimize, differential_evolution
from skopt import gp_minimize
from skopt.space import Real
from skopt.utils import use_named_args
from deap import base, creator, tools, algorithms
import random
from pyswarm import pso
from sko.PSO import PSO as MultiPSO 


class AspenController:
    def __init__(self, filepath, visible=False, timeout=120, debug=False):
        pythoncom.CoInitialize()
        self.filepath = filepath
        self.visible = visible
        self.timeout = timeout
        self.debug = debug
        self.aspen = None
        self._connect()

    def _connect(self):
        try:
            if self.debug:
                print(f"[INFO] Opening Aspen file: {self.filepath}")
            self.aspen = win32com.client.Dispatch("Apwn.Document")
            self.aspen.InitFromArchive2(self.filepath)
            self.aspen.Visible = self.visible
            self.aspen.SuppressDialogs = True
            if self.debug:
                print("[INFO] Aspen file loaded successfully.")
        except Exception as e:
            raise RuntimeError(f"[ERROR] Failed to connect to Aspen: {e}")

    def is_alive(self):
        try:
            return self.aspen and self.aspen.Engine.IsRunning is not None
        except:
            return False

    def ensure_alive(self):
        if not self.is_alive():
            if self.debug:
                print("[WARNING] Aspen instance lost. Reconnecting...")
            self._connect()

    def set(self, path, value):
        self.ensure_alive()
        node = self.aspen.Tree.FindNode(path)
        if node is None:
            raise ValueError(f"[ERROR] Invalid Aspen path: {path}")
        node.Value = value
        if self.debug:
            print(f"[INFO] Set {path} = {value}")

    def get(self, path):
        self.ensure_alive()
        node = self.aspen.Tree.FindNode(path)
        if node is None:
            raise ValueError(f"[ERROR] Invalid Aspen path: {path}")
        value = node.Value
        if self.debug:
            print(f"[INFO] Read {path} = {value}")
        return value

    def run(self):
        self.ensure_alive()
        if self.debug:
            print("[INFO] Starting Aspen simulation...")
        self.aspen.Engine.Run2()
        start_time = time.time()
        while self.aspen.Engine.IsRunning:
            if time.time() - start_time > self.timeout:
                raise TimeoutError("[ERROR] Aspen simulation timed out.")
            time.sleep(1)
        if self.debug:
            print("[INFO] Aspen simulation completed.")

    def close(self):
        if self.aspen:
            try:
                self.aspen.Close(False)
                if self.debug:
                    print("[INFO] Aspen closed.")
            except:
                pass


class SequentialRunner:
    def __init__(self, sim_path, input_vars, output_vars, timeout=120, debug=False):
        self.sim_path = sim_path
        self.input_vars = input_vars
        self.output_vars = output_vars
        self.timeout = timeout
        self.debug = debug
        self.controller = AspenController(sim_path, debug=debug, timeout=timeout)

    def run_batch(self, combinations):
        results = []
        for combo in tqdm(combinations, desc="Running simulations", disable=False):
            input_dict = {path: val for (name, (path, vals)), val in zip(self.input_vars.items(), combo)}
            try:
                for path, value in input_dict.items():
                    self.controller.set(path, value)
                self.controller.run()
                output_dict = {name: self.controller.get(path) for name, path in self.output_vars.items()}
                results.append({**input_dict, **output_dict})
                if self.debug:
                    print("[DEBUG] Simulation successful")
            except Exception as e:
                if self.debug:
                    print(f"[ERROR] Simulation failed: {e}")
                results.append({**input_dict, **{k: None for k in self.output_vars}})
        self.controller.close()
        return results


class ParallelSweep:
    def __init__(self, sim_path, input_vars, output_vars, parallelism=1, timeout=120, debug=False):
        self.sim_path = sim_path
        self.input_vars = input_vars
        self.output_vars = output_vars
        self.timeout = timeout
        self.debug = debug
        self.parallelism = max(1, min(parallelism, cpu_count()))

    def _worker(self, chunk):
        runner = SequentialRunner(
            sim_path=self.sim_path,
            input_vars=self.input_vars,
            output_vars=self.output_vars,
            timeout=self.timeout,
            debug=self.debug
        )
        return runner.run_batch(chunk)

    def run(self):
        param_values = [vals for (_, vals) in self.input_vars.values()]
        # Detect if all param lists are the same length and treat as row-wise samples
        if all(len(v) == len(param_values[0]) for v in param_values) and len(param_values[0]) > 1:
            # Assume inputs are structured row-wise samples, not Cartesian sweep
            all_combinations = list(zip(*param_values))
            print(f"[DEBUG] Detected structured sample input. Using {len(all_combinations)} direct samples.")
        else:
            # Use Cartesian product (parameter sweep)
            all_combinations = list(itertools.product(*param_values))
            print(f"[DEBUG] Detected sweep mode. Generated {len(all_combinations)} combinations.")

        chunks = [all_combinations[i::self.parallelism] for i in range(self.parallelism)]

        print(f"[INFO] Running {len(all_combinations)} combinations using {self.parallelism} Aspen instance(s)")
        if self.parallelism == 1:
            return pd.DataFrame(self._worker(chunks[0]))

        with Pool(processes=self.parallelism) as pool:
            results = pool.map(self._worker, chunks)

        flat_results = [item for sublist in results for item in sublist]
        return pd.DataFrame(flat_results)


class Sweep:
    def __init__(self, sim_path, parallelism=1, timeout=120, debug=False):
        self.sim_path = sim_path
        self.input_vars = {}  # name: (aspen_path, values)
        self.output_vars = {}  # name: aspen_path
        self.parallelism = parallelism
        self.timeout = timeout
        self.debug = debug

    def vary(self, name, path, values):
        self.input_vars[name] = (path, values)
        if self.debug:
            print(f"[DEBUG] Added input variation: {name} = {path} over {len(values)} values")

    def track(self, name, path):
        self.output_vars[name] = path
        if self.debug:
            print(f"[DEBUG] Tracking output: {name} = {path}")

    def run(self, save_csv=False, filename="sweep_results.csv"):
        
        df = ParallelSweep(
            sim_path=self.sim_path,
            input_vars=self.input_vars,
            output_vars=self.output_vars,
            timeout=self.timeout,
            debug=self.debug,
            parallelism=self.parallelism
        ).run()
        
        if save_csv:
            try:
                df.to_csv(filename, index=False)
                if self.debug:
                    print(f"[INFO] Results saved to {filename}")
            except Exception as e:
                print(f"[ERROR] Failed to save CSV: {e}")

        return df
    
    @staticmethod
    def plot_surface(X, Y, input_names=None, output_name="Output", method_name="", interpolate=False, grid_resolution=50):
        """
        Plot a 3D surface of the output response for two input variables.

        Parameters:
            X (np.ndarray): Input samples (N x 2)
            Y (np.ndarray): Output values (N,) or (N x 1)
            input_names (list): Names of input variables [x_name, y_name]
            output_name (str): Label for the output variable
            method_name (str): Optional title suffix (e.g., 'Saltelli', 'Grid')
            interpolate (bool): Whether to interpolate using griddata for smoother surfaces
            grid_resolution (int): Resolution of the interpolation grid
        """
        import matplotlib.pyplot as plt
        from mpl_toolkits.mplot3d import Axes3D
        from scipy.interpolate import griddata
        import numpy as np

        if X.shape[1] != 2:
            raise ValueError("Surface plot requires exactly 2 input variables.")

        x, y = X[:, 0], X[:, 1]
        z = Y.ravel() if len(Y.shape) > 1 else Y

        fig = plt.figure(figsize=(10, 7))
        ax = fig.add_subplot(111, projection='3d')

        if interpolate:
            xi = np.linspace(x.min(), x.max(), grid_resolution)
            yi = np.linspace(y.min(), y.max(), grid_resolution)
            xi, yi = np.meshgrid(xi, yi)
            zi = griddata((x, y), z, (xi, yi), method='cubic')

            ax.plot_surface(xi, yi, zi, cmap='viridis', alpha=0.9, edgecolor='none')
        else:
            ax.plot_trisurf(x, y, z, cmap='viridis', edgecolor='none', alpha=0.9)

        ax.set_xlabel(input_names[0] if input_names else "X1")
        ax.set_ylabel(input_names[1] if input_names else "X2")
        ax.set_zlabel(output_name)
        ax.set_title(f"Response Surface ({method_name})" if method_name else "Response Surface")
        plt.tight_layout()
        plt.show()    



class Sensitivity:
    def __init__(self, sim_path, N=1000, parallelism=1, debug=True, timeout=120):
        self.sim_path = sim_path
        self.N = N
        self.parallelism = parallelism
        self.debug = debug
        self.timeout = timeout
        self.input_vars = {}  # name -> {'path': str, 'bounds': (low, high)}
        self.output_vars = {}  # name -> path
        self.X = None
        self.Y = None
        self.sampling_method = None

    def vary(self, name, path, bounds):
        self.input_vars[name] = {'path': path, 'bounds': bounds}

    def track(self, name, path):
        self.output_vars[name] = path

    def get_problem(self):
        return {
            'num_vars': len(self.input_vars),
            'names': list(self.input_vars.keys()),
            'bounds': [self.input_vars[k]['bounds'] for k in self.input_vars]
        }

    def sample(self, method='saltelli', calc_second_order=True, plot_sample_space=True):
        problem = self.get_problem()
        D = problem['num_vars']
        self.sampling_method = method.lower()

        if self.sampling_method == 'saltelli':
            self.X = saltelli.sample(problem, self.N, calc_second_order=calc_second_order)
        elif self.sampling_method == 'morris':
            self.X = morris.sample(problem, N=self.N, num_levels=4)
        elif self.sampling_method == 'fast':
            self.X = fast_sampler.sample(problem, self.N)
        elif self.sampling_method == 'lhs':
            self.X = latin.sample(problem, self.N)
        elif self.sampling_method == 'sobol_sequence':
            self.X = QmcSobol(d=D, scramble=False).random(self.N)
        elif self.sampling_method == 'halton':
            self.X = Halton(d=D, scramble=False).random(self.N)
        elif self.sampling_method == 'random':
            self.X = np.random.rand(self.N, D)
        elif self.sampling_method == 'grid':
            resolution = int(np.sqrt(self.N)) if D == 2 else int(round(self.N ** (1 / D)))
            grids = [np.linspace(*self.input_vars[name]['bounds'], resolution) for name in problem['names']]
            mesh = np.meshgrid(*grids)
            self.X = np.column_stack([m.flatten() for m in mesh])
            if self.debug:
                print(f"[INFO] Generated {self.X.shape[0]} grid samples ({resolution} per axis in {D}D space)")
        else:
            raise ValueError(f"Unsupported sampling method: {method}")

        if self.debug:
            print(f"[INFO] Generated {len(self.X)} samples using {self.sampling_method}")

        if plot_sample_space:
            dim = self.X.shape[1]
            if dim == 2:
                plt.figure(figsize=(7, 5))
                plt.scatter(self.X[:, 0], self.X[:, 1], alpha=0.6, edgecolor='k')
                plt.xlabel(problem['names'][0])
                plt.ylabel(problem['names'][1])
                plt.title(f"{self.sampling_method.capitalize()} Sample Space (2D)")
                plt.grid(True)
                plt.tight_layout()
                plt.show()
            elif dim == 3:
                from mpl_toolkits.mplot3d import Axes3D
                fig = plt.figure(figsize=(8, 6))
                ax = fig.add_subplot(111, projection='3d')
                ax.scatter(self.X[:, 0], self.X[:, 1], self.X[:, 2], alpha=0.6)
                ax.set_xlabel(problem['names'][0])
                ax.set_ylabel(problem['names'][1])
                ax.set_zlabel(problem['names'][2])
                ax.set_title(f"{self.sampling_method.capitalize()} Sample Space (3D)")
                plt.tight_layout()
                plt.show()
            else:
                print(f"[INFO] Sample space plotting is only supported for 2 or 3 input variables (got {dim}).")

        return self.X



    def run(self):
        input_var_paths = {name: (self.input_vars[name]['path'], []) for name in self.input_vars}
        for row in self.X:
            for i, name in enumerate(input_var_paths):
                input_var_paths[name][1].append(row[i])
        if self.debug:
            print(f"[DEBUG] Preparing structured inputs from X of shape: {self.X.shape}")

        input_vars_structured = {name: (path, values) for name, (path, values) in input_var_paths.items()}

        runner = ParallelSweep(
            sim_path=self.sim_path,
            input_vars=input_vars_structured,
            output_vars=self.output_vars,
            timeout=self.timeout,
            debug=self.debug,
            parallelism=self.parallelism
        )
        if self.debug:
            print(f"[DEBUG] Starting simulations with {len(self.X)} samples...")

        df = runner.run()
        self.Y = df[[k for k in self.output_vars]].values
        return self.Y

    def analyze(self, method='sobol', output_index=0):
        if self.Y is None:
            raise RuntimeError("Call run() before analyze().")

        problem = self.get_problem()
        method = method.lower()
        compatible = {
            'sobol': ['saltelli'],
            'morris': ['morris'],
            'fast': ['fast'],
            'delta': 'any'
        }
        
        if method not in compatible:
            raise ValueError(f"Unsupported analysis method: {method}")
        if compatible[method] != 'any' and self.sampling_method not in compatible[method]:
            raise ValueError(f"{method} analysis requires {compatible[method]} sampling.")

        Y_col = self.Y[:, output_index]

        if self.debug:
            print(f"[INFO] Analyzing output {output_index} using {method}")

        if method == 'sobol':
            return sobol.analyze(problem, Y_col, calc_second_order=True)
        elif method == 'morris':
            return morris_analyze.analyze(problem, self.X, Y_col)
        elif method == 'fast':
            return fast.analyze(problem, Y_col)
        elif method == 'delta':
            return delta.analyze(problem, self.X, Y_col)
    
    @staticmethod
    def plot_surface(X, Y, input_names=None, output_name="Output", method_name="", interpolate=False, grid_resolution=50):
        """
        Plot a 3D surface of the output response for two input variables.

        Parameters:
            X (np.ndarray): Input samples (N x 2)
            Y (np.ndarray): Output values (N,) or (N x 1)
            input_names (list): Names of input variables [x_name, y_name]
            output_name (str): Label for the output variable
            method_name (str): Optional title suffix (e.g., 'Saltelli', 'Grid')
            interpolate (bool): Whether to interpolate using griddata for smoother surfaces
            grid_resolution (int): Resolution of the interpolation grid
        """
        import matplotlib.pyplot as plt
        from mpl_toolkits.mplot3d import Axes3D
        from scipy.interpolate import griddata
        import numpy as np

        if X.shape[1] != 2:
            raise ValueError("Surface plot requires exactly 2 input variables.")

        x, y = X[:, 0], X[:, 1]
        z = Y.ravel() if len(Y.shape) > 1 else Y

        fig = plt.figure(figsize=(10, 7))
        ax = fig.add_subplot(111, projection='3d')

        if interpolate:
            xi = np.linspace(x.min(), x.max(), grid_resolution)
            yi = np.linspace(y.min(), y.max(), grid_resolution)
            xi, yi = np.meshgrid(xi, yi)
            zi = griddata((x, y), z, (xi, yi), method='cubic')

            ax.plot_surface(xi, yi, zi, cmap='viridis', alpha=0.9, edgecolor='none')
        else:
            ax.plot_trisurf(x, y, z, cmap='viridis', edgecolor='none', alpha=0.9)

        ax.set_xlabel(input_names[0] if input_names else "X1")
        ax.set_ylabel(input_names[1] if input_names else "X2")
        ax.set_zlabel(output_name)
        ax.set_title(f"Response Surface ({method_name})" if method_name else "Response Surface")
        plt.tight_layout()
        plt.show()

    @staticmethod
    def plot_morris_trajectories(X, input_names=None, title="Morris Sample Trajectories"):
        """
        Plot Morris sampling trajectories for 2D input space.

        Parameters:
            X (np.ndarray): Morris samples (N * (D + 1), D)
            input_names (list): Optional names of input variables [x, y]
            title (str): Plot title
        """
        N, D = X.shape
        if D != 2:
            raise ValueError("This plot only supports 2D input space (D=2).")

        # Determine number of trajectories
        num_trajectories = N // (D + 1)

        plt.figure(figsize=(8, 6))
        
        # Plot each trajectory
        for i in range(num_trajectories):
            start = i * (D + 1)
            end = (i + 1) * (D + 1)
            traj = X[start:end]
            plt.plot(traj[:, 0], traj[:, 1], '-o', alpha=0.6, linewidth=1.2, markersize=4)

        x_label = input_names[0] if input_names else "X1"
        y_label = input_names[1] if input_names else "X2"

        plt.xlabel(x_label)
        plt.ylabel(y_label)
        plt.title(title)
        plt.grid(True)
        plt.tight_layout()
        plt.show()
    
    @staticmethod
    def plot_sobol_bars(sobol_result, parameter_names=None, figsize=(10, 6)):
        """
        Plot first-order and total-order Sobol sensitivity indices.

        Parameters:
            sobol_result (dict): Output from SALib's sobol.analyze()
            parameter_names (list): Optional list of parameter names
            figsize (tuple): Size of the matplotlib figure
        """
        if parameter_names is None:
            if 'names' in sobol_result:
                parameter_names = sobol_result['names']
            else:
                raise ValueError("Parameter names must be provided explicitly.")


        S1 = sobol_result['S1']
        S1_conf = sobol_result['S1_conf']
        ST = sobol_result['ST']
        ST_conf = sobol_result['ST_conf']

        x = np.arange(len(parameter_names))
        width = 0.35

        fig, ax = plt.subplots(figsize=figsize)
        ax.bar(x - width / 2, S1, width, yerr=S1_conf, capsize=5, label='First-order', color='skyblue')
        ax.bar(x + width / 2, ST, width, yerr=ST_conf, capsize=5, label='Total-order', color='salmon')

        ax.set_ylabel('Sobol Index')
        ax.set_title('Sobol Sensitivity Indices')
        ax.set_xticks(x)
        ax.set_xticklabels(parameter_names, rotation=45, ha='right')
        ax.legend()
        ax.grid(True)
        plt.tight_layout()
        plt.show()

    
    @staticmethod
    def plot_sobol_second_order_heatmap(sobol_result, parameter_names=None, figsize=(8, 6), cmap="coolwarm"):
        """
        Plot a heatmap of second-order Sobol indices.

        Parameters:
            sobol_result (dict): Output from SALib's sobol.analyze()
            parameter_names (list): Optional list of parameter names
            figsize (tuple): Size of the matplotlib figure
            cmap (str): Colormap for the heatmap
        """
        if parameter_names is None:
            parameter_names = sobol_result['names']

        S2 = sobol_result['S2']
        n = len(parameter_names)
        matrix = np.zeros((n, n))

        for i in range(n):
            for j in range(n):
                if i != j:
                    matrix[i, j] = S2[i, j]

        plt.figure(figsize=figsize)
        sns.heatmap(matrix, xticklabels=parameter_names, yticklabels=parameter_names,
                    cmap=cmap, annot=True, fmt=".2f", square=True, cbar_kws={'label': 'Second-order Sobol Index'})
        plt.title('Second-order Sobol Sensitivity Heatmap')
        plt.tight_layout()
        plt.show()

    
    @staticmethod
    def plot_violin_outputs(X, Y, parameter_names, output_name="Output", figsize=(10, 6)):
        """
        Create violin plots showing distribution of output Y across input bins.

        Parameters:
            X (np.ndarray): Sampled input matrix (N x D)
            Y (np.ndarray): Output values (N,)
            parameter_names (list): Names of input parameters
            output_name (str): Label for Y
            figsize (tuple): Size of the figure
        """
        n_bins = 10
        df_list = []

        for i, name in enumerate(parameter_names):
            x = X[:, i]
            bins = np.linspace(np.min(x), np.max(x), n_bins + 1)
            labels = [f"{round(bins[j], 2)}–{round(bins[j+1], 2)}" for j in range(n_bins)]

            digitized = np.digitize(x, bins) - 1
            digitized = np.clip(digitized, 0, n_bins - 1)

            for j in range(n_bins):
                bin_vals = Y[digitized == j]
                for val in bin_vals:
                    df_list.append({'Parameter': name, 'Bin': labels[j], output_name: val})

        df = pd.DataFrame(df_list)

        plt.figure(figsize=figsize)
        sns.violinplot(data=df, x='Parameter', y=output_name, hue='Bin', split=True, palette='Set2')
        plt.title('Output Distribution Across Input Parameter Bins')
        plt.xticks(rotation=45)
        plt.tight_layout()
        plt.show()


class AspenControlPanel:
    def __init__(self, title="Aspen Control Panel"):
        self.root = tk.Tk()
        self.root.title(title)
        self.root.geometry("350x200")
        self.root.resizable(False, False)

        self.label = tk.Label(self.root, text="Aspen Controller", font=("Arial", 14))
        self.label.pack(pady=5)

        self.status_label = tk.Label(self.root, text="Checking instances...", font=("Arial", 11))
        self.status_label.pack()

        self.kill_button = tk.Button(
            self.root,
            text="Kill Aspen Instances",
            font=("Arial", 12),
            bg="red",
            fg="white",
            command=self.kill_aspen_instances
        )
        self.kill_button.pack(pady=10)

        self.refresh_button = tk.Button(
            self.root,
            text="Refresh Status",
            font=("Arial", 10),
            command=self.update_status
        )
        self.refresh_button.pack()

        self.browse_button = tk.Button(
            self.root,
            text="Open Aspen File",
            font=("Arial", 10),
            command=self.open_aspen_file
        )
        self.browse_button.pack(pady=5)

        self.update_status_periodically()

    def update_status(self):
        count = 0
        for proc in psutil.process_iter(attrs=['name']):
            if proc.info['name'] == "AspenPlus.exe":
                count += 1
        self.status_label.config(text=f"Aspen instances running: {count}")

    def update_status_periodically(self):
        self.update_status()
        self.root.after(5000, self.update_status_periodically)

    def kill_aspen_instances(self):
        try:
            os.system("taskkill /f /im AspenPlus.exe >nul 2>&1")
            messagebox.showinfo("Success", "All Aspen Plus instances terminated.")
            self.update_status()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to terminate Aspen: {e}")

    def open_aspen_file(self):
        filepath = tk.filedialog.askopenfilename(
            title="Open Aspen File",
            filetypes=[("Aspen Backup Files", "*.bkp"), ("All files", "*.*")]
        )
        if filepath:
            try:
                subprocess.Popen(["AspenPlus.exe", filepath], shell=True)
                messagebox.showinfo("Success", f"Opened {filepath}")
            except Exception as e:
                messagebox.showerror("Error", f"Could not open file: {e}")

    def run(self):
        self.root.mainloop()



class Optimize:
    def __init__(self, sim_path, timeout=120, debug=True, normalize=False, parallelism=1):
        self.sim_path = sim_path
        self.timeout = timeout
        self.debug = debug
        self.normalize = normalize
        self.parallelism = parallelism
        self.input_vars = {}  # name: (path, bounds)
        self.output_vars = {}  # name: path
        self.objective_funcs = []
        self.controller = AspenController(filepath=self.sim_path, timeout=self.timeout, debug=self.debug)

    def vary(self, name, path, bounds):
        self.input_vars[name] = (path, bounds)

    def track(self, name, path):
        self.output_vars[name] = path

    def add_objective(self, func):
        self.objective_funcs.append(func)

    def normalize_inputs(self, x):
        bounds = [b for (_, b) in self.input_vars.values()]
        return [(val - low) / (high - low) for val, (low, high) in zip(x, bounds)]

    def unnormalize_inputs(self, x_norm):
        bounds = [b for (_, b) in self.input_vars.values()]
        return [low + xi * (high - low) for xi, (low, high) in zip(x_norm, bounds)]

    def _evaluate(self, x):
        if self.normalize:
            x = self.unnormalize_inputs(x)

        try:
            for i, (name, (path, _)) in enumerate(self.input_vars.items()):
                self.controller.set(path, x[i])
            self.controller.run()

            outputs = {}
            for name, path in self.output_vars.items():
                try:
                    outputs[name] = self.controller.get(path)
                except Exception as e:
                    outputs[name] = np.nan
                    if self.debug:
                        print(f"[WARN] Could not read {name}: {e}")

            if len(self.objective_funcs) == 1:
                return self.objective_funcs[0](outputs)
            return tuple(func(outputs) for func in self.objective_funcs)

        except Exception as e:
            if self.debug:
                print(f"[ERROR] Evaluation failed: {e}")
            if len(self.objective_funcs) == 1:
                return 1e8
            return tuple([1e8] * len(self.objective_funcs))

    def close(self):
        self.controller.close()

    def run_single_objective(self, method='Nelder-Mead', **kwargs):
        bounds = [b for (_, b) in self.input_vars.values()]
        if self.normalize:
            bounds = [(0, 1)] * len(bounds)
            x0 = [0.5] * len(bounds)
        else:
            x0 = [(b[0] + b[1]) / 2 for b in bounds]

        method = method.lower()

        if method in ['nelder-mead', 'powell', 'slsqp', 'bfgs', 'l-bfgs-b', 'tnc']:
            return minimize(
                self._evaluate,
                x0,
                method=method,
                bounds=bounds if method in ['l-bfgs-b', 'slsqp', 'tnc'] else None
            )

        elif method == 'differential_evolution':
            return differential_evolution(self._evaluate, bounds, disp=self.debug, **kwargs)

        elif method == 'bayesian':
            if len(self.objective_funcs) != 1:
                raise ValueError("Bayesian optimization requires a single-objective function.")
            space = [Real(*b, name=name) for name, (_, b) in self.input_vars.items()]
            if self.normalize:
                space = [Real(0, 1, name=name) for name in self.input_vars]

            @use_named_args(space)
            def objective(**kwargs2):
                x = [kwargs2[name] for name in self.input_vars]
                return self._evaluate(x)

            return gp_minimize(
                objective,
                space,
                n_calls=kwargs.get('n_calls', 30),
                acq_func=kwargs.get('acq_func', 'EI'),
                verbose=self.debug
            )

        elif method == 'pso':
            from pyswarm import pso
            lb = [b[0] for (_, b) in self.input_vars.values()]
            ub = [b[1] for (_, b) in self.input_vars.values()]
            xopt, fopt = pso(self._evaluate, lb, ub, debug=self.debug, **kwargs)
            return xopt, fopt

        else:
            raise ValueError(f"Unsupported optimization method: {method}")

    def run_multi_objective(self, method='nsga2', ngen=50, pop_size=40, cxpb=0.7, mutpb=0.2, **kwargs):
        if len(self.objective_funcs) < 2:
            raise ValueError("Multi-objective optimization requires at least two objectives.")

        bounds = [b for (_, b) in self.input_vars.values()]
        if self.normalize:
            bounds = [(0, 1)] * len(bounds)

        if method == 'nsga2':
            dim = len(bounds)

            if not hasattr(creator, "FitnessMulti"):
                creator.create("FitnessMulti", base.Fitness, weights=tuple([-1.0] * len(self.objective_funcs)))
            if not hasattr(creator, "Individual"):
                creator.create("Individual", list, fitness=creator.FitnessMulti)

            toolbox = base.Toolbox()
            toolbox.register("individual", tools.initIterate, creator.Individual,
                             lambda: [random.uniform(*b) for b in bounds])
            toolbox.register("population", tools.initRepeat, list, toolbox.individual)

            toolbox.register("evaluate", self._evaluate)
            toolbox.register("mate", tools.cxBlend, alpha=0.5)
            toolbox.register("mutate", tools.mutPolynomialBounded,
                             eta=20.0, low=[b[0] for b in bounds], up=[b[1] for b in bounds], indpb=0.1)
            toolbox.register("select", tools.selNSGA2)

            pop = toolbox.population(n=pop_size)
            algorithms.eaMuPlusLambda(pop, toolbox, mu=pop_size, lambda_=pop_size, cxpb=cxpb, mutpb=mutpb,
                                      ngen=ngen, verbose=self.debug)

            return tools.sortNondominated(pop, len(pop), first_front_only=True)[0]

        elif method == 'pso':
            from mopso import MOPSO

            lb = [b[0] for (_, b) in self.input_vars.values()]
            ub = [b[1] for (_, b) in self.input_vars.values()]

            def obj_func(x):
                return list(self._evaluate(x))

            mopso = MOPSO(obj_func=obj_func, lb=lb, ub=ub, pop_size=pop_size, max_gen=ngen, dim=len(bounds), debug=self.debug)
            pareto_front = mopso.run()
            return pareto_front

        else:
            raise ValueError(f"Unsupported multi-objective method: {method}")

    @staticmethod
    def plot_pareto_front(population, labels=None, title="Pareto Front"):
        import matplotlib.pyplot as plt

        objectives = np.array([ind.fitness.values for ind in population])
        if objectives.shape[1] != 2:
            raise ValueError("Plotting only supported for 2-objective optimization.")

        plt.figure(figsize=(8, 6))
        plt.scatter(objectives[:, 0], objectives[:, 1], c='blue', edgecolors='k', alpha=0.7)
        plt.xlabel(labels[0] if labels else "Objective 1")
        plt.ylabel(labels[1] if labels else "Objective 2")
        plt.title(title)
        plt.grid(True)
        plt.tight_layout()
        plt.show()

    @staticmethod
    def save_pareto_to_csv(population, input_names, filename="pareto_solutions.csv"):
        data = []
        for ind in population:
            inputs = ind
            objectives = ind.fitness.values
            row = dict(zip(input_names, inputs))
            for i, obj in enumerate(objectives):
                row[f"Objective_{i+1}"] = obj
            data.append(row)

        df = pd.DataFrame(data)
        df.to_csv(filename, index=False)
        print(f"[INFO] Pareto solutions saved to {filename}")

    @staticmethod
    def summarize_single_result(result, input_names):
        x = result[0] if isinstance(result, tuple) else result.x
        fx = result[1] if isinstance(result, tuple) else result.fun
        print("[INFO] Best solution found:")
        for name, val in zip(input_names, x):
            print(f" - {name}: {val:.6f}")
        print(f"Objective value: {fx:.6f}")
