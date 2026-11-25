import java.util.*;

/** VATIA - Análisis de fallas por componentes simétricas (puerto de FaultAnalysisScreen).
 *  Requiere Java 17+.  javac FallasSim.java && java FallasSim */
public class FallasSim {
    record C(double r, double i) {
        C add(C o) { return new C(r + o.r, i + o.i); }
        C sub(C o) { return new C(r - o.r, i - o.i); }
        C mul(C o) { return new C(r * o.r - i * o.i, r * o.i + i * o.r); }
        C div(C o) { double d = o.r * o.r + o.i * o.i; return new C((r * o.r + i * o.i) / d, (i * o.r - r * o.i) / d); }
        C neg() { return new C(-r, -i); }
        C k(double s) { return new C(r * s, i * s); }
        double abs() { return Math.hypot(r, i); }
        double deg() { return Math.toDegrees(Math.atan2(i, r)); }
        public String toString() { return String.format("%.4f ∠ %.2f°", abs(), deg()); }
    }
    static final C Z0C = new C(0, 0), ONE = new C(1, 0);

    record Linea(int a, int b, C z1, C z0, double bsh) {}
    record Trafo(int a, int b, C z, String cfg) {}
    record Shunt(int bus, double q) {}
    record Carga(int bus, double p, double q) {}
    record Gen(int bus, C z1, C z2, C z0) {}

    static C[][] ybus(int n, int seq, double S, List<Linea> L, List<Trafo> T, List<Shunt> ind,
                      List<Shunt> cap, List<Carga> cargas, List<Gen> G) {
        C[][] Y = new C[n][n];
        for (C[] r : Y) Arrays.fill(r, Z0C);
        for (Linea l : L) {
            C z = seq == 0 ? l.z0 : l.z1;
            if (z.abs() < 1e-9) continue;
            C y = ONE.div(z);
            Y[l.a][l.a] = Y[l.a][l.a].add(y); Y[l.b][l.b] = Y[l.b][l.b].add(y);
            Y[l.a][l.b] = Y[l.a][l.b].sub(y); Y[l.b][l.a] = Y[l.b][l.a].sub(y);
            if (seq != 0) {
                Y[l.a][l.a] = Y[l.a][l.a].add(new C(0, l.bsh)); Y[l.b][l.b] = Y[l.b][l.b].add(new C(0, l.bsh));
            }
        }
        for (Trafo t : T) {
            if (t.z.abs() < 1e-9) continue;
            C y = ONE.div(t.z);
            boolean full = seq != 0 || t.cfg.contains("YNyn");
            if (full) {
                Y[t.a][t.a] = Y[t.a][t.a].add(y); Y[t.b][t.b] = Y[t.b][t.b].add(y);
                Y[t.a][t.b] = Y[t.a][t.b].sub(y); Y[t.b][t.a] = Y[t.b][t.a].sub(y);
            } else if (t.cfg.contains("YNd")) Y[t.a][t.a] = Y[t.a][t.a].add(y);
            else if (t.cfg.contains("Dyn")) Y[t.b][t.b] = Y[t.b][t.b].add(y);
        }
        for (Shunt s : ind) Y[s.bus][s.bus] = Y[s.bus][s.bus].add(new C(0, -s.q / S));
        if (seq != 0) {
            for (Shunt s : cap) Y[s.bus][s.bus] = Y[s.bus][s.bus].add(new C(0, s.q / S));
            for (Carga c : cargas) Y[c.bus][c.bus] = Y[c.bus][c.bus].add(new C(c.p / S, -c.q / S));
        }
        for (Gen g : G) {
            C z = seq == 0 ? g.z0 : g.z1;       // igual que el Python: sec. 2 usa Z1 del generador
            if (z.abs() < 1e-9) continue;
            Y[g.bus][g.bus] = Y[g.bus][g.bus].add(ONE.div(z));
        }
        return Y;
    }

    static C[][] inv(C[][] M) {
        int n = M.length;
        C[][] A = new C[n][2 * n];
        for (int i = 0; i < n; i++) for (int j = 0; j < n; j++) { A[i][j] = M[i][j]; A[i][n + j] = i == j ? ONE : Z0C; }
        for (int c = 0; c < n; c++) {
            int p = c;
            for (int r = c + 1; r < n; r++) if (A[r][c].abs() > A[p][c].abs()) p = r;
            if (A[p][c].abs() < 1e-12) throw new ArithmeticException("Ybus singular");
            C[] t = A[c]; A[c] = A[p]; A[p] = t;
            C d = A[c][c];
            for (int j = 0; j < 2 * n; j++) A[c][j] = A[c][j].div(d);
            for (int r = 0; r < n; r++) if (r != c) {
                C f = A[r][c];
                for (int j = 0; j < 2 * n; j++) A[r][j] = A[r][j].sub(f.mul(A[c][j]));
            }
        }
        C[][] R = new C[n][n];
        for (int i = 0; i < n; i++) System.arraycopy(A[i], n, R[i], 0, n);
        return R;
    }

    /** Devuelve {I0, I1, I2} en la falla.  tipo: SLG, LL, LLG, LLL */
    static C[] resolver(String t, C V, C Z0, C Z1, C Z2, C ZF, C Zg) {
        C three = new C(3, 0);
        switch (t) {
            case "SLG": { C d = Z0.add(Z1).add(Z2).add(three.mul(ZF)).add(three.mul(Zg)); C I = V.div(d); return new C[]{I, I, I}; }
            case "LLL": { C I = V.div(Z1.add(ZF)); return new C[]{Z0C, I, Z0C}; }
            case "LL":  { C I = V.div(Z1.add(Z2).add(ZF)); return new C[]{Z0C, I, I.neg()}; }
            default: {
                C Za = Z0.add(three.mul(Zg)).add(ZF), Zb = Z2.add(ZF);
                C d02 = Z0.add(Z2).add(ZF.k(2)).add(three.mul(Zg));
                C ZT = Za.mul(Zb).div(Za.add(Zb));
                C I1 = V.div(Z1.add(ZF).add(ZT));
                return new C[]{I1.mul(Z2.add(ZF)).div(d02).neg(), I1, I1.mul(Za).div(d02).neg()};
            }
        }
    }

    public static void main(String[] args) {
        int n = 2; double S = 100, VB = 220;
        List<Linea> L = List.of(new Linea(0, 1, new C(0, .1), new C(0, .3), 0));
        List<Gen> G = List.of(new Gen(0, new C(0, .1), new C(0, .05), new C(0, .05)));
        List<Trafo> T = List.of(); List<Shunt> ind = List.of(), cap = List.of(); List<Carga> cg = List.of();
        C[][][] Zb = new C[3][][];
        for (int s = 0; s < 3; s++) Zb[s] = inv(ybus(n, s, S, L, T, ind, cap, cg, G));
        int k = 1; C Vp = ONE, ZF = Z0C, Zg = Z0C;
        double Ibase = S / (Math.sqrt(3) * VB);
        C a = new C(-0.5, Math.sqrt(3) / 2), a2 = a.mul(a);
        for (String tipo : new String[]{"LLL", "SLG", "LL", "LLG"}) {
            C[] I = resolver(tipo, Vp, Zb[0][k][k], Zb[1][k][k], Zb[2][k][k], ZF, Zg);
            C Ia = I[0].add(I[1]).add(I[2]);
            C Ib = I[0].add(a2.mul(I[1])).add(a.mul(I[2]));
            C Ic = I[0].add(a.mul(I[1])).add(a2.mul(I[2]));
            System.out.printf("%-4s |Ia|=%.4f |Ib|=%.4f |Ic|=%.4f p.u.  (Ia = %.3f kA)%n",
                tipo, Ia.abs(), Ib.abs(), Ic.abs(), Ia.abs() * Ibase);
        }
    }
}
