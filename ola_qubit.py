from qiskit import QuantumCircuit, transpile
from qiskit_aer import AerSimulator
from qiskit.visualization import plot_histogram
import matplotlib.pyplot as plt

qc = QuantumCircuit(1, 1)
qc.h(0)
qc.measure(0, 0)

sim = AerSimulator()
tqc = transpile(qc, sim)   
job = sim.run(tqc, shots=1000)  
result = job.result()
counts = result.get_counts()
print(counts)
plot_histogram(counts)
plt.show()