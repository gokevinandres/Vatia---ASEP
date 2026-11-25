import java.util.*;

/** VATIA - Despacho económico DC con congestión (puerto de EconomicDispatchCore).
 *  Incluye: merit order, flujo DC, ISF, redespacho iterativo y precios nodales.
 *  (No incluye la enumeración de alternativas de redespacho.)  Java 17+. */
public class DespachoED {
    record Linea(int a, int b, double x, double cap) {}
    record Gen(int bus, double cap, double mc) {}

    final int n; final double[] loads; final List<Linea> lines; final List<Gen> gens; final int slack;
    final int[] ns; final double[][] Br; final double[][] ISF;

    static double[] solve(double[][] A, double[] b) {
        int n = b.length; double[][] M = new double[n][n + 1];
        for (int i = 0; i < n; i++) { System.arraycopy(A[i], 0, M[i], 0, n); M[i][n] = b[i]; }
        for (int c = 0; c < n; c++) {
            int p = c;
            for (int r = c + 1; r < n; r++) if (Math.abs(M[r][c]) > Math.abs(M[p][c])) p = r;
            if (Math.abs(M[p][c]) < 1e-12) throw new ArithmeticException("Matriz B reducida singular");
            double[] t = M[c]; M[c] = M[p]; M[p] = t;
            for (int r = c + 1; r < n; r++) { double k = M[r][c] / M[c][c]; for (int j = c; j <= n; j++) M[r][j] -= k * M[c][j]; }
        }
        double[] x = new double[n];
        for (int i = n - 1; i >= 0; i--) { double s = M[i][n]; for (int j = i + 1; j < n; j++) s -= M[i][j] * x[j]; x[i] = s / M[i][i]; }
        return x;
    }

    DespachoED(int n, double[] loads, List<Linea> lines, List<Gen> gens, int slack) {
        this.n = n; this.loads = loads; this.lines = lines; this.gens = gens; this.slack = slack;
        double[][] Bf = new double[n][n];
        for (Linea L : lines) {
            if (Math.abs(L.x) < 1e-12) continue;
            double y = 1 / L.x;
            Bf[L.a][L.a] += y; Bf[L.b][L.b] += y; Bf[L.a][L.b] -= y; Bf[L.b][L.a] -= y;
        }
        ns = new int[n - 1];
        for (int i = 0, q = 0; i < n; i++) if (i != slack) ns[q++] = i;
        int m = n - 1; Br = new double[m][m];
        for (int i = 0; i < m; i++) for (int j = 0; j < m; j++) Br[i][j] = Bf[ns[i]][ns[j]];
        double[][] Xe = new double[n][n];
        for (int j = 0; j < m; j++) {
            double[] e = new double[m]; e[j] = 1;
            double[] col = solve(Br, e);
            for (int i = 0; i < m; i++) Xe[ns[i]][ns[j]] = col[i];
        }
        ISF = new double[lines.size()][n];
        for (int k = 0; k < lines.size(); k++) {
            Linea L = lines.get(k);
            if (Math.abs(L.x) < 1e-12) continue;
            for (int i = 0; i < n; i++) ISF[k][i] = (Xe[L.a][i] - Xe[L.b][i]) / L.x;
        }
    }

    double[] flows(double[] Pg) {
        double[] Pb = new double[n];
        for (int g = 0; g < gens.size(); g++) Pb[gens.get(g).bus] += Pg[g];
        for (int i = 0; i < n; i++) Pb[i] -= loads[i];
        double[] pr = new double[n - 1];
        for (int i = 0; i < n - 1; i++) pr[i] = Pb[ns[i]];
        double[] th = new double[n], tr;
        try { tr = solve(Br, pr); } catch (ArithmeticException e) { return null; }
        for (int i = 0; i < n - 1; i++) th[ns[i]] = tr[i];
        double[] F = new double[lines.size()];
        for (int k = 0; k < F.length; k++) { Linea L = lines.get(k); if (Math.abs(L.x) >= 1e-12) F[k] = (th[L.a] - th[L.b]) / L.x; }
        return F;
    }

    double[] merit() {
        double rem = 0; for (double l : loads) rem += l;
        Integer[] ord = new Integer[gens.size()];
        for (int i = 0; i < ord.length; i++) ord[i] = i;
        Arrays.sort(ord, Comparator.comparingDouble(g -> gens.get(g).mc));
        double[] P = new double[gens.size()];
        for (int g : ord) { double d = Math.min(gens.get(g).cap, rem); P[g] = d; rem -= d; if (rem <= 1e-9) break; }
        return P;
    }

    /** Redespacho iterativo. Devuelve P final; it[0] = iteraciones, it[1] = 1 si convergió. */
    double[] redespachar(double[] P0, int[] it) {
        double[] P = P0.clone(); int G = gens.size(); boolean conv = false; int i;
        for (i = 0; i < 30; i++) {
            double[] fl = flows(P);
            if (fl == null) { it[0] = i; it[1] = 0; return P; }
            int wk = -1; double we = 0;
            for (int k = 0; k < lines.size(); k++) { double ex = Math.abs(fl[k]) - lines.get(k).cap; if (ex > we) { we = ex; wk = k; } }
            if (wk < 0 || we < 1e-3) { conv = true; break; }
            double sf = fl[wk] > 0 ? 1 : -1; double[] best = null;
            for (int u = 0; u < G; u++) {
                double au = gens.get(u).cap - P[u]; if (au < 1e-6) continue;
                for (int d = 0; d < G; d++) {
                    if (u == d || P[d] < 1e-6) continue;
                    double df = sf * (ISF[wk][gens.get(u).bus] - ISF[wk][gens.get(d).bus]);
                    if (df >= -1e-9) continue;
                    double del = Math.min(we / (-df), Math.min(au, P[d]));
                    if (del < 1e-9) continue;
                    double c = (gens.get(u).mc - gens.get(d).mc) * del;
                    if (best == null || c < best[2]) best = new double[]{u, d, c, del};
                }
            }
            if (best == null) break;
            P[(int) best[0]] += best[3]; P[(int) best[1]] -= best[3];
        }
        it[0] = Math.min(i + 1, 30); it[1] = conv ? 1 : 0;
        return P;
    }

    double precio(int k, double[] P, double[] fl) {
        int G = gens.size(); List<Integer> cg = new ArrayList<>();
        for (int l = 0; l < lines.size(); l++) if (Math.abs(fl[l]) >= lines.get(l).cap - 1e-3) cg.add(l);
        double maxMc = 0; for (Gen g : gens) maxMc = Math.max(maxMc, g.mc);
        if (cg.isEmpty()) {
            for (int g = 0; g < G; g++) if (P[g] > 1e-6 && P[g] < gens.get(g).cap - 1e-6) return gens.get(g).mc;
            return maxMc;
        }
        Double best = null;
        for (int u = 0; u < G; u++) {
            int bu = gens.get(u).bus; double au = gens.get(u).cap - P[u]; if (au < 1 - 1e-6) continue;
            boolean ok = true;
            for (int l : cg) { double sf = fl[l] > 0 ? 1 : -1; if (sf * (ISF[l][bu] - ISF[l][k]) > 1e-9) { ok = false; break; } }
            if (ok && (best == null || gens.get(u).mc < best)) best = gens.get(u).mc;
        }
        for (int u = 0; u < G; u++) {
            int bu = gens.get(u).bus; double au = gens.get(u).cap - P[u]; if (au < 1 - 1e-6) continue;
            for (int d = 0; d < G; d++) {
                if (u == d) continue;
                int bd = gens.get(d).bus; double ad = P[d]; if (ad < 1e-6) continue;
                double lo = 1, hi = Math.min(au, ad + 1); if (lo > hi + 1e-9) continue;
                boolean fe = true;
                for (int l : cg) {
                    double sf = fl[l] > 0 ? 1 : -1;
                    double A = sf * (ISF[l][bu] - ISF[l][bd]), B = sf * (ISF[l][bd] - ISF[l][k]);
                    if (Math.abs(A) < 1e-9) { if (B > 1e-9) { fe = false; break; } }
                    else if (A > 0) hi = Math.min(hi, -B / A); else lo = Math.max(lo, -B / A);
                }
                if (!fe || lo > hi + 1e-9) continue;
                double ao = gens.get(u).mc > gens.get(d).mc ? lo : hi;
                double c = ao * gens.get(u).mc - (ao - 1) * gens.get(d).mc;
                if (c < 0) c = Math.max(gens.get(u).mc, gens.get(d).mc);
                if (best == null || c < best) best = c;
            }
        }
        return best == null ? maxMc : Math.max(best, 0);
    }

    public static void main(String[] args) {
        List<Linea> L = List.of(new Linea(0, 1, .1, 50), new Linea(1, 2, .1, 60), new Linea(0, 2, .1, 100));
        List<Gen> G = List.of(new Gen(0, 200, 20), new Gen(2, 150, 35));
        DespachoED ed = new DespachoED(3, new double[]{0, 50, 100}, L, G, 0);
        double[] PED = ed.merit(), fED = ed.flows(PED);
        int[] it = new int[2];
        double[] Pf = ed.redespachar(PED, it), ff = ed.flows(Pf);
        System.out.println("P_ED      = " + Arrays.toString(PED));
        System.out.println("Flujos ED = " + Arrays.toString(fED));
        System.out.println("P_final   = " + Arrays.toString(Pf) + "  (iter " + it[0] + ", convergió=" + (it[1] == 1) + ")");
        System.out.println("Flujos fin= " + Arrays.toString(ff));
        for (int k = 0; k < 3; k++) System.out.printf("Precio nodal %d = %.2f $/MWh%n", k + 1, ed.precio(k, Pf, ff));
    }
}
