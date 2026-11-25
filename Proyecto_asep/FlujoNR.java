import java.util.*;

/** VATIA - Flujos de potencia Newton-Raphson (puerto del módulo AppFlujoPotencia).
 *  Compilar: javac FlujoNR.java   Ejecutar: java FlujoNR */
public class FlujoNR {
    static final int PQ = 1, PV = 2, SLACK = 3;

    static class Nodo {
        int tipo; double Pd, Qd, Pg, Qg, V, thDeg, V0 = 1, th0Deg = 0;
        Nodo(int tipo, double Pd, double Qd, double Pg, double Qg, double V, double thDeg) {
            this.tipo = tipo; this.Pd = Pd; this.Qd = Qd; this.Pg = Pg; this.Qg = Qg; this.V = V; this.thDeg = thDeg;
        }
    }
    static class Linea {
        int a, b; double G, B, sh;   // nodos base 1, y = G + jB, sh = y_shunt/2 (p.u.)
        Linea(int a, int b, double G, double B, double sh) { this.a = a - 1; this.b = b - 1; this.G = G; this.B = B; this.sh = sh; }
    }

    static double[] solve(double[][] A, double[] b) {
        int n = b.length;
        double[][] M = new double[n][n + 1];
        for (int i = 0; i < n; i++) { System.arraycopy(A[i], 0, M[i], 0, n); M[i][n] = b[i]; }
        for (int c = 0; c < n; c++) {
            int p = c;
            for (int r = c + 1; r < n; r++) if (Math.abs(M[r][c]) > Math.abs(M[p][c])) p = r;
            if (Math.abs(M[p][c]) < 1e-14) throw new ArithmeticException("Jacobiano singular");
            double[] t = M[c]; M[c] = M[p]; M[p] = t;
            for (int r = c + 1; r < n; r++) {
                double k = M[r][c] / M[c][c];
                for (int j = c; j <= n; j++) M[r][j] -= k * M[c][j];
            }
        }
        double[] x = new double[n];
        for (int i = n - 1; i >= 0; i--) {
            double s = M[i][n];
            for (int j = i + 1; j < n; j++) s -= M[i][j] * x[j];
            x[i] = s / M[i][i];
        }
        return x;
    }

    public static void ejecutar(List<Nodo> nodos, List<Linea> lineas, double sBase, double vBaseKv, int nIter) {
        int n = nodos.size();
        double[][] G = new double[n][n], B = new double[n][n];
        for (Linea l : lineas) {
            G[l.a][l.a] += l.G; B[l.a][l.a] += l.B + l.sh;
            G[l.b][l.b] += l.G; B[l.b][l.b] += l.B + l.sh;
            G[l.a][l.b] -= l.G; B[l.a][l.b] -= l.B;
            G[l.b][l.a] -= l.G; B[l.b][l.a] -= l.B;
        }
        double[] V = new double[n], th = new double[n], Pe = new double[n], Qe = new double[n];
        List<Integer> ns = new ArrayList<>(), pq = new ArrayList<>();
        for (int i = 0; i < n; i++) {
            Nodo d = nodos.get(i);
            Pe[i] = (d.Pg - d.Pd) / sBase; Qe[i] = (d.Qg - d.Qd) / sBase;
            if (d.tipo == PQ) { V[i] = d.V0; th[i] = Math.toRadians(d.th0Deg); }
            else { V[i] = d.V; th[i] = Math.toRadians(d.thDeg); }
            if (d.tipo != SLACK) ns.add(i);
            if (d.tipo == PQ) pq.add(i);
        }
        int m = ns.size() + pq.size();
        double[] P = new double[n], Q = new double[n];

        for (int k = 0; k < nIter; k++) {
            for (int i = 0; i < n; i++) {
                double p = 0, q = 0;
                for (int j = 0; j < n; j++) {
                    double d = th[i] - th[j];
                    p += V[j] * (G[i][j] * Math.cos(d) + B[i][j] * Math.sin(d));
                    q += V[j] * (G[i][j] * Math.sin(d) - B[i][j] * Math.cos(d));
                }
                P[i] = V[i] * p; Q[i] = V[i] * q;
            }
            double[] mis = new double[m];
            int r = 0;
            for (int i : ns) mis[r++] = Pe[i] - P[i];
            for (int i : pq) mis[r++] = Qe[i] - Q[i];

            double[][] J = new double[m][m];
            r = 0;
            for (int i : ns) {
                int c = 0;
                for (int j : ns) {
                    double d = th[i] - th[j];
                    J[r][c++] = (i == j) ? -Q[i] - B[i][i] * V[i] * V[i]
                        : V[i] * V[j] * (G[i][j] * Math.sin(d) - B[i][j] * Math.cos(d));
                }
                for (int kk : pq) {
                    double d = th[i] - th[kk];
                    J[r][c++] = (i == kk) ? (Math.abs(V[i]) < 1e-12 ? 0 : P[i] / V[i] + G[i][i] * V[i])
                        : V[i] * (G[i][kk] * Math.cos(d) + B[i][kk] * Math.sin(d));
                }
                r++;
            }
            for (int i : pq) {
                int c = 0;
                for (int j : ns) {
                    double d = th[i] - th[j];
                    J[r][c++] = (i == j) ? P[i] - G[i][i] * V[i] * V[i]
                        : -V[i] * V[j] * (G[i][j] * Math.cos(d) + B[i][j] * Math.sin(d));
                }
                for (int kk : pq) {
                    double d = th[i] - th[kk];
                    J[r][c++] = (i == kk) ? (Math.abs(V[i]) < 1e-12 ? 0 : Q[i] / V[i] - B[i][i] * V[i])
                        : V[i] * (G[i][kk] * Math.sin(d) - B[i][kk] * Math.cos(d));
                }
                r++;
            }
            double[] dx = solve(J, mis);

            double err = 0;
            for (double v : mis) err = Math.max(err, Math.abs(v));
            System.out.printf("%n=== Iteración %d  (máx |mismatch| = %.3e) ===%n", k, err);
            System.out.println("Nodo  Tipo   V(p.u.)   θ(rad)     P_calc      Q_calc");
            for (int i = 0; i < n; i++)
                System.out.printf("%3d   %-5s %8.4f %9.5f %11.6f %11.6f%n", i + 1,
                    new String[]{"", "PQ", "PV", "Slack"}[nodos.get(i).tipo], V[i], th[i], P[i], Q[i]);

            r = 0;
            for (int i : ns) th[i] += dx[r++];
            for (int i : pq) V[i] += dx[r++];
        }

        // Resultados finales con el estado de la última iteración calculada
        System.out.println("\n=== RESULTADOS FINALES ===");
        double pg = 0, qg = 0, pd = 0, qd = 0;
        System.out.println("Nodo  Tipo    V(kV)    θ(°)     Pg(MW)    Qg(MVAr)   Pd(MW)   Qd(MVAr)");
        for (int i = 0; i < n; i++) {
            Nodo d = nodos.get(i);
            double Pg = d.Pd + P[i] * sBase, Qg = d.Qd + Q[i] * sBase;
            pg += Pg; qg += Qg; pd += d.Pd; qd += d.Qd;
            System.out.printf("%3d   %-5s %8.2f %9.4f %9.2f %10.2f %9.2f %9.2f%n", i + 1,
                new String[]{"", "PQ", "PV", "Slack"}[d.tipo], V[i] * vBaseKv, Math.toDegrees(th[i]), Pg, Qg, d.Pd, d.Qd);
        }
        System.out.printf("Generación: %.2f MW / %.2f MVAr | Demanda: %.2f MW / %.2f MVAr | Pérdidas: %.2f MW / %.2f MVAr%n",
            pg, qg, pd, qd, pg - pd, qg - qd);
    }

    public static void main(String[] args) {
        // Ejemplo de 3 nodos (misma estructura de datos que las tablas de la app)
        List<Nodo> nodos = List.of(
            new Nodo(SLACK, 0, 0, 0, 0, 1.0, 0),
            new Nodo(PQ, 50, 20, 0, 0, 1.0, 0),
            new Nodo(PQ, 30, 10, 0, 0, 1.0, 0));
        List<Linea> lineas = List.of(
            new Linea(1, 2, 1, -10, 0), new Linea(2, 3, 1, -10, 0), new Linea(1, 3, 1, -10, 0));
        ejecutar(nodos, lineas, 100.0, 220.0, 4);   // iteraciones 0 a 3
    }
}
