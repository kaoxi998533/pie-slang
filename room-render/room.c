/*
 * room.c — 一个小型蒙特卡洛路径追踪器，渲染一间有阳光照进来的客厅。
 *
 * 全部几何与纹理都是程序化生成的（不读取任何外部图片/模型）：
 *   - 房间墙体/窗框/书架用解析轴对齐盒子求交
 *   - 沙发、茶几、植物、窗帘等用 SDF（有符号距离场）球步进求交，可做圆角
 *   - 光源：太阳（小角半径锥体采样）+ 透过窗洞的天空（面积采样）
 *   - 漫反射 + 菲涅尔清漆层（地板/陶瓷/漆面）+ 金属
 *   - 下一事件估计（NEE）+ 俄罗斯轮盘赌
 *
 * 输出：原始 float32 缓冲（颜色、首次命中反照率、法线、深度），交给 post.py 做
 * 降噪、泛光、色调映射并存成 PNG。
 *
 * 编译：gcc -O3 -march=native -fopenmp room.c -o room -lm
 * 运行：./room <宽> <高> <每像素采样数> <输出.bin>
 */
#include <math.h>
#include <stdint.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <time.h>
#ifdef _OPENMP
#include <omp.h>
#endif

#ifndef M_PI
#define M_PI 3.14159265358979323846
#endif

/* ------------------------------------------------------------------ 向量 */
typedef struct { double x, y, z; } V;
static inline V v3(double x, double y, double z) { V r = {x, y, z}; return r; }
static inline V add(V a, V b) { return v3(a.x + b.x, a.y + b.y, a.z + b.z); }
static inline V sub(V a, V b) { return v3(a.x - b.x, a.y - b.y, a.z - b.z); }
static inline V mul(V a, double s) { return v3(a.x * s, a.y * s, a.z * s); }
static inline V vmul(V a, V b) { return v3(a.x * b.x, a.y * b.y, a.z * b.z); }
static inline double dot(V a, V b) { return a.x * b.x + a.y * b.y + a.z * b.z; }
static inline V cross(V a, V b) { return v3(a.y * b.z - a.z * b.y, a.z * b.x - a.x * b.z, a.x * b.y - a.y * b.x); }
static inline double len(V a) { return sqrt(dot(a, a)); }
static inline V norm(V a) { return mul(a, 1.0 / len(a)); }
static inline V vabs(V a) { return v3(fabs(a.x), fabs(a.y), fabs(a.z)); }
static inline V vmax0(V a) { return v3(fmax(a.x, 0), fmax(a.y, 0), fmax(a.z, 0)); }
static inline double maxc(V a) { return fmax(a.x, fmax(a.y, a.z)); }
static inline double clampd(double x, double a, double b) { return x < a ? a : (x > b ? b : x); }
static inline double mixd(double a, double b, double t) { return a + (b - a) * t; }
static inline V mixv(V a, V b, double t) { return add(a, mul(sub(b, a), t)); }
static inline double smoothstep(double e0, double e1, double x) {
    double t = clampd((x - e0) / (e1 - e0), 0, 1);
    return t * t * (3 - 2 * t);
}
static inline double fract(double x) { return x - floor(x); }
static inline V reflect(V d, V n) { return sub(d, mul(n, 2 * dot(d, n))); }

/* ------------------------------------------------------------------ 随机数 (PCG32) */
typedef struct { uint64_t s; } Rng;
static inline uint32_t pcg(Rng *r) {
    uint64_t old = r->s;
    r->s = old * 6364136223846793005ULL + 1442695040888963407ULL;
    uint32_t xs = (uint32_t)(((old >> 18u) ^ old) >> 27u);
    uint32_t rot = (uint32_t)(old >> 59u);
    return (xs >> rot) | (xs << ((-rot) & 31));
}
static inline double rnd(Rng *r) { return pcg(r) * (1.0 / 4294967296.0); }

/* ------------------------------------------------------------------ 噪声 */
static inline uint32_t ihash(uint32_t x) {
    x ^= x >> 16; x *= 0x7feb352dU; x ^= x >> 15; x *= 0x846ca68bU; x ^= x >> 16;
    return x;
}
static inline double h3(int x, int y, int z) {
    return ihash((uint32_t)x * 73856093U ^ ihash((uint32_t)y * 19349663U ^ ihash((uint32_t)z * 83492791U))) /
           4294967296.0;
}
static double vnoise(V p) {
    int ix = (int)floor(p.x), iy = (int)floor(p.y), iz = (int)floor(p.z);
    double fx = p.x - ix, fy = p.y - iy, fz = p.z - iz;
    fx = fx * fx * (3 - 2 * fx); fy = fy * fy * (3 - 2 * fy); fz = fz * fz * (3 - 2 * fz);
    double a = mixd(h3(ix, iy, iz), h3(ix + 1, iy, iz), fx);
    double b = mixd(h3(ix, iy + 1, iz), h3(ix + 1, iy + 1, iz), fx);
    double c = mixd(h3(ix, iy, iz + 1), h3(ix + 1, iy, iz + 1), fx);
    double d = mixd(h3(ix, iy + 1, iz + 1), h3(ix + 1, iy + 1, iz + 1), fx);
    return mixd(mixd(a, b, fy), mixd(c, d, fy), fz);
}
static double fbm(V p, int oct) {
    double s = 0, a = 0.5, tot = 0;
    for (int i = 0; i < oct; i++) {
        s += a * vnoise(p);
        tot += a;
        p = v3(p.x * 2.03 + 1.7, p.y * 2.03 - 3.1, p.z * 2.03 + 5.3);
        a *= 0.5;
    }
    return s / tot;
}

/* ------------------------------------------------------------------ 场景物体 */
enum { T_BOX = 0, T_RBOX, T_CYL, T_ELL, T_CAP, T_CUP, T_TUBE, T_CURT };
enum {
    M_WALL, M_CEIL, M_FLOOR, M_TRIM, M_SOFA, M_PILLOW_MUSTARD, M_PILLOW_CREAM, M_WALNUT, M_RUG,
    M_BOOK_FLAT, M_BOOK_UP, M_CERAMIC, M_BLACK, M_BRASS, M_SHADE, M_CURTAIN, M_POT, M_SOIL, M_LEAF,
    M_STEM, M_ART, M_OAK
};

typedef struct {
    int type, mat;
    V c, ax, ay, az; /* 中心与局部坐标轴（世界->局部 的行向量） */
    V h;             /* 半尺寸 / 半径 */
    double r, ph, freq;
    V a, b;          /* 胶囊端点 */
    V bmin, bmax;    /* 包围盒 */
    V col;           /* 每个物体自己的颜色（书、叶子） */
    int id;
} Obj;

#define MAXO 2048
static Obj objs[MAXO];
static int nobj = 0;

typedef struct { V bmin, bmax; int s, e; } Group;
static Group groups[64];
static int ngroup = 0;

static Obj *newobj(int type, int mat) {
    if (nobj >= MAXO) { fprintf(stderr, "too many objects\n"); exit(1); }
    Obj *o = &objs[nobj++];
    memset(o, 0, sizeof *o);
    o->type = type; o->mat = mat;
    o->ax = v3(1, 0, 0); o->ay = v3(0, 1, 0); o->az = v3(0, 0, 1);
    o->col = v3(0.5, 0.5, 0.5);
    return o;
}
static void begin_group(void) { groups[ngroup].s = nobj; }
static void end_group(void) {
    Group *g = &groups[ngroup++];
    g->e = nobj;
    g->bmin = v3(1e9, 1e9, 1e9); g->bmax = v3(-1e9, -1e9, -1e9);
    for (int i = g->s; i < g->e; i++) {
        g->bmin = v3(fmin(g->bmin.x, objs[i].bmin.x), fmin(g->bmin.y, objs[i].bmin.y), fmin(g->bmin.z, objs[i].bmin.z));
        g->bmax = v3(fmax(g->bmax.x, objs[i].bmax.x), fmax(g->bmax.y, objs[i].bmax.y), fmax(g->bmax.z, objs[i].bmax.z));
    }
}

/* 旋转 R = Ry(yaw) * Rx(pitch) * Rz(roll)；局部->世界 的列向量即世界->局部 的行向量 */
static void set_rot(Obj *o, double yaw, double pitch, double roll) {
    double cy = cos(yaw), sy = sin(yaw), cx = cos(pitch), sx = sin(pitch), cz = cos(roll), sz = sin(roll);
    double Ry[3][3] = {{cy, 0, sy}, {0, 1, 0}, {-sy, 0, cy}};
    double Rx[3][3] = {{1, 0, 0}, {0, cx, -sx}, {0, sx, cx}};
    double Rz[3][3] = {{cz, -sz, 0}, {sz, cz, 0}, {0, 0, 1}};
    double A[3][3], R[3][3];
    for (int i = 0; i < 3; i++)
        for (int j = 0; j < 3; j++) {
            A[i][j] = 0;
            for (int k = 0; k < 3; k++) A[i][j] += Ry[i][k] * Rx[k][j];
        }
    for (int i = 0; i < 3; i++)
        for (int j = 0; j < 3; j++) {
            R[i][j] = 0;
            for (int k = 0; k < 3; k++) R[i][j] += A[i][k] * Rz[k][j];
        }
    o->ax = v3(R[0][0], R[1][0], R[2][0]);
    o->ay = v3(R[0][1], R[1][1], R[2][1]);
    o->az = v3(R[0][2], R[1][2], R[2][2]);
}
static void finish_bounds(Obj *o, V H) {
    V e;
    e.x = fabs(o->ax.x) * H.x + fabs(o->ay.x) * H.y + fabs(o->az.x) * H.z;
    e.y = fabs(o->ax.y) * H.x + fabs(o->ay.y) * H.y + fabs(o->az.y) * H.z;
    e.z = fabs(o->ax.z) * H.x + fabs(o->ay.z) * H.y + fabs(o->az.z) * H.z;
    double m = 0.003;
    o->bmin = sub(sub(o->c, e), v3(m, m, m));
    o->bmax = add(add(o->c, e), v3(m, m, m));
}

static Obj *box(V a, V b, int mat) {
    Obj *o = newobj(T_BOX, mat);
    o->bmin = a; o->bmax = b;
    o->c = mul(add(a, b), 0.5); o->h = mul(sub(b, a), 0.5);
    return o;
}
static Obj *rbox(V c, V h, double r, int mat, double yaw, double pitch, double roll) {
    Obj *o = newobj(T_RBOX, mat);
    o->c = c; o->h = h; o->r = r;
    set_rot(o, yaw, pitch, roll);
    finish_bounds(o, h);
    return o;
}
static Obj *cyl(V c, double ra, double hh, double rr, int mat) {
    Obj *o = newobj(T_CYL, mat);
    o->c = c; o->h = v3(ra, hh, ra); o->r = rr;
    finish_bounds(o, o->h);
    return o;
}
static Obj *cap(V a, V b, double r, int mat) {
    Obj *o = newobj(T_CAP, mat);
    o->a = a; o->b = b; o->r = r;
    o->c = mul(add(a, b), 0.5);
    o->bmin = v3(fmin(a.x, b.x) - r - 0.002, fmin(a.y, b.y) - r - 0.002, fmin(a.z, b.z) - r - 0.002);
    o->bmax = v3(fmax(a.x, b.x) + r + 0.002, fmax(a.y, b.y) + r + 0.002, fmax(a.z, b.z) + r + 0.002);
    return o;
}
static Obj *cup(V c, double ra, double hh, double t, int mat) {
    Obj *o = newobj(T_CUP, mat);
    o->c = c; o->h = v3(ra, hh, ra); o->r = t;
    finish_bounds(o, o->h);
    return o;
}
static Obj *tube(V c, double ra, double hh, double t, int mat) {
    Obj *o = newobj(T_TUBE, mat);
    o->c = c; o->h = v3(ra, hh, ra); o->r = t;
    finish_bounds(o, v3(ra + t, hh, ra + t));
    return o;
}
/* 窗帘：贴墙（法线沿 x）的正弦褶皱薄片 */
static Obj *curtain(V c, double hy, double hz, double amp, double freq, double th, double phase, int mat) {
    Obj *o = newobj(T_CURT, mat);
    o->c = c; o->h = v3(amp, hy, hz); o->r = th; o->freq = freq; o->ph = phase;
    finish_bounds(o, v3(amp * 1.35 + th, hy, hz));
    return o;
}
/* 叶片：弯曲压扁的椭球；f = 叶片长轴方向 */
static Obj *leaf(V c, V f, double twist, V radii, double bend, V col) {
    Obj *o = newobj(T_ELL, M_LEAF);
    o->c = c; o->h = radii; o->ph = bend; o->col = col;
    V up = fabs(f.y) > 0.95 ? v3(1, 0, 0) : v3(0, 1, 0);
    V x = norm(cross(up, f));
    V y = cross(f, x);
    double ct = cos(twist), st = sin(twist);
    o->ax = add(mul(x, ct), mul(y, st));
    o->ay = sub(mul(y, ct), mul(x, st));
    o->az = f;
    finish_bounds(o, v3(radii.x, radii.y + bend * 0.05 + 0.01, radii.z));
    return o;
}

/* ------------------------------------------------------------------ SDF */
static inline double sd_rbox(V q, V b, double r) {
    V qq = sub(vabs(q), sub(b, v3(r, r, r)));
    return len(vmax0(qq)) + fmin(fmax(qq.x, fmax(qq.y, qq.z)), 0) - r;
}
static inline double sd_cyl(V q, double ra, double hh, double rr) {
    double dx = sqrt(q.x * q.x + q.z * q.z) - ra + rr;
    double dy = fabs(q.y) - hh + rr;
    double ox = fmax(dx, 0), oy = fmax(dy, 0);
    return fmin(fmax(dx, dy), 0) + sqrt(ox * ox + oy * oy) - rr;
}
static inline double sd_ell(V q, V r) {
    double k0 = len(v3(q.x / r.x, q.y / r.y, q.z / r.z));
    double k1 = len(v3(q.x / (r.x * r.x), q.y / (r.y * r.y), q.z / (r.z * r.z)));
    if (k1 < 1e-12) return -fmin(r.x, fmin(r.y, r.z));
    return k0 * (k0 - 1) / k1;
}
static inline double sd_cap(V p, V a, V b, double r) {
    V pa = sub(p, a), ba = sub(b, a);
    double hh = clampd(dot(pa, ba) / dot(ba, ba), 0, 1);
    return len(sub(pa, mul(ba, hh))) - r;
}
static inline V to_local(const Obj *o, V p) {
    V d = sub(p, o->c);
    return v3(dot(d, o->ax), dot(d, o->ay), dot(d, o->az));
}
static double obj_sdf(const Obj *o, V p) {
    if (o->type == T_CAP) return sd_cap(p, o->a, o->b, o->r);
    V q = to_local(o, p);
    switch (o->type) {
    case T_RBOX: return sd_rbox(q, o->h, o->r);
    case T_CYL: return sd_cyl(q, o->h.x, o->h.y, o->r);
    case T_ELL: {
        /* 叶片：横向微卷、纵向下垂；尖端更宽 */
        q.y += o->ph * (3.0 * q.x * q.x - 1.2 * q.z * q.z);
        V r = o->h;
        r.x *= 1.0 + 0.25 * clampd(q.z / o->h.z, -1, 1);
        return 0.55 * sd_ell(q, r);
    }
    case T_CUP: {
        double outer = sd_cyl(q, o->h.x, o->h.y, 0.004);
        double inner = sd_cyl(v3(q.x, q.y - o->r, q.z), o->h.x - o->r, o->h.y, 0.002);
        return fmax(outer, -inner);
    }
    case T_TUBE:
        return fmax(fabs(sqrt(q.x * q.x + q.z * q.z) - o->h.x) - o->r, fabs(q.y) - o->h.y);
    case T_CURT: {
        /* 下摆处褶皱稍大，随高度有轻微相位漂移 */
        double t = (o->h.y - q.y) / (2 * o->h.y);
        double amp = o->h.x * (0.8 + 0.5 * t);
        double dx = q.x - amp * sin(q.z * o->freq + o->ph + 0.6 * sin(q.y * 2.1 + o->ph));
        double d1 = (fabs(dx) - o->r) * 0.42;
        double d2 = fmax(fabs(q.y) - o->h.y, fabs(q.z) - o->h.z);
        return fmax(d1, d2);
    }
    }
    return 1e9;
}
static V sdf_normal(const Obj *o, V p) {
    const double e = 2e-4;
    V k1 = v3(1, -1, -1), k2 = v3(-1, -1, 1), k3 = v3(-1, 1, -1), k4 = v3(1, 1, 1);
    V n = add(add(mul(k1, obj_sdf(o, add(p, mul(k1, e)))), mul(k2, obj_sdf(o, add(p, mul(k2, e))))),
              add(mul(k3, obj_sdf(o, add(p, mul(k3, e)))), mul(k4, obj_sdf(o, add(p, mul(k4, e))))));
    double l = len(n);
    return l > 0 ? mul(n, 1.0 / l) : v3(0, 1, 0);
}

/* ------------------------------------------------------------------ 求交 */
static inline int slab(V o, V inv, V bmin, V bmax, double *t0, double *t1) {
    double tx1 = (bmin.x - o.x) * inv.x, tx2 = (bmax.x - o.x) * inv.x;
    double tmin = fmin(tx1, tx2), tmax = fmax(tx1, tx2);
    double ty1 = (bmin.y - o.y) * inv.y, ty2 = (bmax.y - o.y) * inv.y;
    tmin = fmax(tmin, fmin(ty1, ty2)); tmax = fmin(tmax, fmax(ty1, ty2));
    double tz1 = (bmin.z - o.z) * inv.z, tz2 = (bmax.z - o.z) * inv.z;
    tmin = fmax(tmin, fmin(tz1, tz2)); tmax = fmin(tmax, fmax(tz1, tz2));
    *t0 = tmin; *t1 = tmax;
    return tmax >= fmax(tmin, 0.0);
}

typedef struct { double t; int obj; V n; } Hit;

static int trace(V o, V d, double tmax, Hit *hit, int shadow) {
    V inv = v3(1.0 / d.x, 1.0 / d.y, 1.0 / d.z);
    double best = tmax;
    int bo = -1;
    for (int g = 0; g < ngroup; g++) {
        double g0, g1;
        if (!slab(o, inv, groups[g].bmin, groups[g].bmax, &g0, &g1) || g0 > best) continue;
        for (int i = groups[g].s; i < groups[g].e; i++) {
            const Obj *ob = &objs[i];
            double t0, t1;
            if (!slab(o, inv, ob->bmin, ob->bmax, &t0, &t1) || t0 > best) continue;
            if (ob->type == T_BOX) {
                if (t0 > 1e-7 && t0 < best) { best = t0; bo = i; }
            } else {
                double t = fmax(t0, 0.0), tend = fmin(t1, best);
                for (int k = 0; k < 320 && t <= tend; k++) {
                    double dist = obj_sdf(ob, add(o, mul(d, t)));
                    if (dist < 1.2e-4) {
                        if (t < best) { best = t; bo = i; }
                        break;
                    }
                    t += dist;
                }
            }
            if (shadow && bo >= 0) return 1;
        }
    }
    if (bo < 0) return 0;
    hit->t = best;
    hit->obj = bo;
    V p = add(o, mul(d, best));
    const Obj *ob = &objs[bo];
    if (ob->type == T_BOX) {
        double dd[6] = {fabs(p.x - ob->bmin.x), fabs(p.x - ob->bmax.x), fabs(p.y - ob->bmin.y),
                        fabs(p.y - ob->bmax.y), fabs(p.z - ob->bmin.z), fabs(p.z - ob->bmax.z)};
        int k = 0;
        for (int j = 1; j < 6; j++)
            if (dd[j] < dd[k]) k = j;
        V ns[6] = {v3(-1, 0, 0), v3(1, 0, 0), v3(0, -1, 0), v3(0, 1, 0), v3(0, 0, -1), v3(0, 0, 1)};
        hit->n = ns[k];
    } else {
        hit->n = sdf_normal(ob, p);
    }
    return 1;
}

/* ------------------------------------------------------------------ 光照环境 */
#define WALL_T 0.20
static const double WIN_Y0 = 0.75, WIN_Y1 = 2.35, WIN_Z0 = 0.55, WIN_Z1 = 2.35;
static V SUN_DIR;   /* 指向太阳 */
static double SUN_COS;
static V SUN_E;     /* 垂直入射照度 */
static V SUN_L;     /* 太阳圆盘辐亮度 */

static V env(V d, int with_sun) {
    V col;
    double az = atan2(d.z, -d.x);
    double tree = 0.05 + 0.16 * fbm(v3(az * 3.0 + 3.1, 0.5, 1.7), 4) + 0.04 * fbm(v3(az * 22, 2.3, 0.1), 3);
    if (d.y < 0) {
        double k = fmin(1, -d.y * 4);
        double g = 0.85 + 0.3 * fbm(v3(az * 30, d.y * 60, 4.2), 3);
        col = mul(mixv(v3(0.62, 0.74, 0.42), v3(0.42, 0.58, 0.24), k), g);
    } else if (d.y < tree) {
        double sh = fbm(v3(az * 40, d.y * 40, 0.3), 3);
        double lit = 0.35 + 0.65 * smoothstep(0.0, tree, d.y);
        col = mul(v3(0.30, 0.42, 0.16), (0.5 + 1.2 * sh) * lit * 1.6);
    } else {
        double t = pow(fmax(d.y, 0), 0.45);
        col = mixv(v3(1.75, 1.85, 1.95), v3(0.45, 0.72, 1.45), t);
        double c = fmax(0, dot(d, SUN_DIR));
        col = add(col, mul(v3(1.0, 0.85, 0.62), 1.6 * pow(c, 10) + 5.0 * pow(c, 160)));
    }
    if (with_sun && dot(d, SUN_DIR) > SUN_COS) col = add(col, SUN_L);
    return col;
}

static void onb(V n, V *t, V *b) {
    double s = n.z >= 0 ? 1.0 : -1.0;
    double a = -1.0 / (s + n.z);
    double bb = n.x * n.y * a;
    *t = v3(1 + s * n.x * n.x * a, s * bb, -s * n.x);
    *b = v3(bb, s + n.y * n.y * a, -n.y);
}
static V sample_lobe(V axis, double e, Rng *r) {
    double u = rnd(r), v = rnd(r);
    double ct = pow(u, 1.0 / (e + 1));
    double st = sqrt(fmax(0, 1 - ct * ct));
    double ph = 2 * M_PI * v;
    V t, b;
    onb(axis, &t, &b);
    return norm(add(add(mul(t, st * cos(ph)), mul(b, st * sin(ph))), mul(axis, ct)));
}
static V sample_cone(V axis, double cosmax, Rng *r) {
    double ct = 1 - rnd(r) * (1 - cosmax);
    double st = sqrt(fmax(0, 1 - ct * ct));
    double ph = 2 * M_PI * rnd(r);
    V t, b;
    onb(axis, &t, &b);
    return norm(add(add(mul(t, st * cos(ph)), mul(b, st * sin(ph))), mul(axis, ct)));
}
static inline double rough_exp(double r) { return 2.0 / (r * r) - 2.0; }

/* 漫反射点的直接光（太阳 + 窗洞天空），返回入射照度（未乘 albedo/pi） */
static V direct(V x, V n, Rng *rng) {
    V Ld = v3(0, 0, 0);
    Hit h;
    V l = sample_cone(SUN_DIR, SUN_COS, rng);
    double c = dot(n, l);
    if (c > 0 && !trace(x, l, 1e30, &h, 1)) Ld = add(Ld, mul(SUN_E, c));

    double wy = WIN_Y1 - WIN_Y0, wz = WIN_Z1 - WIN_Z0;
    V q = v3(-WALL_T, WIN_Y0 + rnd(rng) * wy, WIN_Z0 + rnd(rng) * wz);
    V w = sub(q, x);
    double d2 = dot(w, w), dist = sqrt(d2);
    V lw = mul(w, 1.0 / dist);
    double cs = dot(n, lw), cw = -lw.x;
    if (cs > 0 && cw > 0 && !trace(x, lw, 1e30, &h, 1))
        Ld = add(Ld, mul(env(lw, 0), cs * cw * wy * wz / d2));
    return Ld;
}

/* ------------------------------------------------------------------ 材质与纹理 */
typedef struct { V alb; double coat, rough; int metal; } Surf;

static V art_color(int id, double u, double v) {
    /* u,v ∈ [0,1]，v 向上 */
    V c;
    double paper = 0.94 + 0.06 * fbm(v3(u * 60, v * 60, id * 7.0), 3);
    if (id == 0) {
        /* 远山落日 */
        c = mixv(v3(0.78, 0.46, 0.30), v3(0.80, 0.72, 0.58), smoothstep(0.35, 1.0, v));
        double sd = len(v3(u - 0.66, v - 0.62, 0));
        if (sd < 0.09) c = v3(0.85, 0.52, 0.28);
        V layers[4] = {v3(0.42, 0.40, 0.42), v3(0.25, 0.27, 0.30), v3(0.12, 0.16, 0.18), v3(0.04, 0.07, 0.08)};
        for (int k = 0; k < 4; k++) {
            double hgt = 0.62 - k * 0.12 + 0.10 * (fbm(v3(u * (2.5 + k), k * 3.7, 1.3), 4) - 0.5) * 2;
            if (v < hgt) c = layers[k];
        }
    } else {
        /* 中古几何抽象 */
        c = v3(0.74, 0.66, 0.55);
        if (len(v3(u - 0.40, v - 0.40, 0)) < 0.27) c = v3(0.52, 0.17, 0.07);
        double ax = fabs(u - 0.70), top = 0.55;
        int arch = (ax < 0.16 && v < top) || (len(v3(u - 0.70, v - top, 0)) < 0.16);
        if (arch && v > 0.12) c = v3(0.20, 0.22, 0.11);
        if (len(v3(u - 0.28, v - 0.78, 0)) < 0.07) c = v3(0.04, 0.04, 0.04);
        if (v < 0.12) c = v3(0.60, 0.48, 0.33);
    }
    return mul(c, paper);
}

static void surface(const Obj *o, V p, V n, Surf *s) {
    s->coat = 0; s->rough = 0.3; s->metal = 0;
    switch (o->mat) {
    case M_WALL: {
        double k = 0.97 + 0.06 * fbm(mul(p, 9.0), 3);
        s->alb = mul(v3(0.70, 0.66, 0.60), k);
        break;
    }
    case M_CEIL: s->alb = v3(0.78, 0.77, 0.74); break;
    case M_TRIM: s->alb = v3(0.76, 0.75, 0.72); s->coat = 0.05; s->rough = 0.25; break;
    case M_FLOOR: {
        const double pw = 0.15, plen = 1.35;
        double fx = p.x / pw;
        int row = (int)floor(fx);
        double u = fx - row;
        double off = h3(row, 7, 1) * plen;
        double fz = (p.z + off) / plen;
        int col = (int)floor(fz);
        double v = fz - col;
        double id1 = h3(row, col, 3), id2 = h3(row, col, 11);
        V base = mixv(v3(0.44, 0.24, 0.11), v3(0.30, 0.15, 0.065), id1);
        base = mul(base, 0.82 + 0.3 * id2);
        double gn = fbm(v3(p.x * 8, p.z * 0.7, id1 * 13), 4);
        double ring = 0.5 + 0.5 * sin(p.x * 260 + gn * 24 + id2 * 40);
        ring = ring * ring * ring;
        base = mul(base, 1 - 0.20 * ring);
        base = mul(base, 0.90 + 0.18 * vnoise(v3(p.x * 420, p.z * 7, row)));
        base = mul(base, 0.90 + 0.2 * fbm(v3(p.x * 1.2, 0, p.z * 1.2), 3));
        double gu = fmin(u, 1 - u) * pw, gv = fmin(v, 1 - v) * plen;
        if (gu < 0.0016 || gv < 0.0016) base = mul(base, 0.25);
        s->alb = base;
        s->coat = 0.045;
        s->rough = 0.09 + 0.10 * fbm(v3(p.x * 3, 0, p.z * 3), 3);
        break;
    }
    case M_SOFA: {
        double k = 0.90 + 0.18 * fbm(mul(p, 22.0), 3);
        s->alb = mul(v3(0.085, 0.12, 0.17), k);
        break;
    }
    case M_PILLOW_MUSTARD: s->alb = mul(v3(0.55, 0.32, 0.06), 0.9 + 0.2 * fbm(mul(p, 40.0), 3)); break;
    case M_PILLOW_CREAM: {
        double b = fbm(mul(p, 90.0), 2);
        s->alb = mul(v3(0.62, 0.57, 0.48), 0.82 + 0.3 * b);
        break;
    }
    case M_WALNUT: {
        V q = to_local(o, p);
        double gn = fbm(v3(q.x * 2.5, q.y * 30, q.z * 30), 4);
        double ring = 0.5 + 0.5 * sin(q.z * 180 + q.y * 120 + gn * 18);
        V base = mixv(v3(0.17, 0.085, 0.042), v3(0.09, 0.045, 0.022), ring * ring);
        s->alb = mul(base, 0.9 + 0.2 * vnoise(v3(q.x * 8, q.z * 300, q.y * 300)));
        s->coat = 0.05; s->rough = 0.12;
        break;
    }
    case M_OAK: {
        double gn = fbm(v3(p.x * 30, p.y * 3, p.z * 30), 4);
        double ring = 0.5 + 0.5 * sin((p.x + p.z) * 200 + gn * 20);
        s->alb = mixv(v3(0.50, 0.32, 0.17), v3(0.38, 0.23, 0.11), ring * ring);
        s->coat = 0.04; s->rough = 0.25;
        break;
    }
    case M_RUG: {
        V q = to_local(o, p);
        double edge = fmin(o->h.x - fabs(q.x), o->h.z - fabs(q.z));
        V cream = v3(0.60, 0.55, 0.45), dark = v3(0.07, 0.07, 0.08), rust = v3(0.36, 0.13, 0.06);
        V c = cream;
        if (edge < 0.07) c = dark;
        else if (edge < 0.10) c = cream;
        else if (edge < 0.115) c = rust;
        else {
            double a = fabs(fract((q.x + q.z) * 2.2) - 0.5), b = fabs(fract((q.x - q.z) * 2.2) - 0.5);
            double line = fmin(a, b);
            if (line < 0.03) c = mixv(cream, dark, 0.8);
            else if (fabs(line - 0.25) < 0.012) c = mixv(cream, rust, 0.6);
        }
        double fib = 0.80 + 0.34 * vnoise(mul(p, 220.0)) + 0.1 * fbm(mul(p, 12.0), 2);
        s->alb = mul(c, fib);
        break;
    }
    case M_BOOK_FLAT: {
        V nl = v3(dot(n, o->ax), dot(n, o->ay), dot(n, o->az));
        V q = to_local(o, p);
        int cover = fabs(nl.y) > 0.6 || nl.x < -0.6;
        if (cover) {
            s->alb = o->col;
            if (fabs(nl.y) > 0.6 && q.x < -o->h.x + 0.02) s->alb = mul(o->col, 0.7);
            s->coat = 0.03; s->rough = 0.35;
        } else {
            double lines = 0.85 + 0.15 * vnoise(v3(0, q.y * 1500, 0));
            s->alb = mul(v3(0.72, 0.68, 0.58), lines);
        }
        break;
    }
    case M_BOOK_UP: {
        if (n.y > 0.6) {
            s->alb = mul(v3(0.70, 0.66, 0.56), 0.9 + 0.1 * vnoise(v3(p.x * 2000, 0, 0)));
        } else {
            double rel = (p.y - o->bmin.y) / (o->bmax.y - o->bmin.y);
            double band = (fabs(rel - 0.85) < 0.03 || fabs(rel - 0.12) < 0.02) ? 0.6 : 1.0;
            s->alb = mul(o->col, band * (0.92 + 0.12 * vnoise(mul(p, 80.0))));
            s->coat = 0.03; s->rough = 0.4;
        }
        break;
    }
    case M_CERAMIC: s->alb = v3(0.74, 0.72, 0.67); s->coat = 0.05; s->rough = 0.05; break;
    case M_BLACK: s->alb = v3(0.025, 0.025, 0.025); s->coat = 0.05; s->rough = 0.3; break;
    case M_BRASS: s->alb = v3(0.75, 0.54, 0.26); s->metal = 1; s->rough = 0.22; break;
    case M_SHADE: s->alb = mul(v3(0.74, 0.69, 0.58), 0.92 + 0.12 * fbm(v3(p.x * 300, p.y * 40, p.z * 300), 2)); break;
    case M_CURTAIN: {
        double st = 0.93 + 0.1 * fbm(v3(p.x * 50, p.y * 3, p.z * 50), 3);
        s->alb = mul(v3(0.66, 0.62, 0.53), st);
        break;
    }
    case M_POT: s->alb = mul(v3(0.46, 0.18, 0.09), 0.92 + 0.12 * fbm(mul(p, 30.0), 3)); break;
    case M_SOIL: s->alb = mul(v3(0.05, 0.035, 0.022), 0.7 + 0.6 * vnoise(mul(p, 300.0))); break;
    case M_LEAF: {
        V q = to_local(o, p);
        double vein = fabs(q.x) < 0.0025 ? 1.6 : 1.0;
        s->alb = mul(o->col, vein * (0.9 + 0.2 * vnoise(mul(p, 120.0))));
        s->coat = 0.05; s->rough = 0.28;
        break;
    }
    case M_STEM: s->alb = v3(0.13, 0.09, 0.05); break;
    case M_ART: {
        V q = to_local(o, p);
        V nl = v3(dot(n, o->ax), dot(n, o->ay), dot(n, o->az));
        V frame = v3(0.025, 0.022, 0.02);
        s->alb = frame; s->coat = 0.04; s->rough = 0.3;
        if (nl.z > 0.6) {
            double bx = o->h.x - fabs(q.x), by = o->h.y - fabs(q.y);
            double b = fmin(bx, by);
            if (b > 0.022) {
                s->coat = 0;
                if (b < 0.075) s->alb = v3(0.80, 0.79, 0.76);
                else {
                    double iw = o->h.x - 0.075, ih = o->h.y - 0.075;
                    s->alb = art_color(o->id, (q.x + iw) / (2 * iw), (q.y + ih) / (2 * ih));
                }
            }
        }
        break;
    }
    default: s->alb = v3(0.5, 0.5, 0.5);
    }
}

/* ------------------------------------------------------------------ 路径追踪 */
typedef struct { V alb, nrm; double depth; int set; } Aov;

static V radiance(V o, V d, Rng *rng, Aov *aov) {
    V L = v3(0, 0, 0), T = v3(1, 1, 1);
    int spec = 1, delta = 1;
    for (int depth = 0; depth < 10; depth++) {
        Hit h;
        if (!trace(o, d, 1e30, &h, 0)) {
            if (spec) {
                V e = env(d, delta);
                if (depth > 0) e = v3(fmin(e.x, 12), fmin(e.y, 12), fmin(e.z, 12));
                L = add(L, vmul(T, e));
            }
            if (!aov->set) { aov->alb = v3(1, 1, 1); aov->nrm = mul(d, -1); aov->depth = 50; aov->set = 1; }
            break;
        }
        V p = add(o, mul(d, h.t));
        V n = h.n;
        if (dot(n, d) > 0) n = mul(n, -1);
        Surf s;
        surface(&objs[h.obj], p, n, &s);
        if (!aov->set) { aov->alb = s.alb; aov->nrm = n; aov->depth = h.t; aov->set = 1; }
        V po = add(p, mul(n, 1e-3));
        double cosi = fmax(0, -dot(d, n));
        V nd;
        if (s.metal) {
            V r = reflect(d, n);
            nd = sample_lobe(r, rough_exp(s.rough), rng);
            if (dot(nd, n) <= 0) break;
            T = vmul(T, s.alb);
            spec = 1; delta = 0;
        } else {
            double F = s.coat > 0 ? s.coat + (1 - s.coat) * pow(1 - cosi, 5) : 0;
            if (rnd(rng) < F) {
                V r = reflect(d, n);
                nd = sample_lobe(r, rough_exp(s.rough), rng);
                if (dot(nd, n) <= 0) break;
                spec = 1; delta = 0;
            } else {
                V Ld = direct(po, n, rng);
                V c = vmul(T, mul(vmul(s.alb, Ld), 1.0 / M_PI));
                if (depth > 0) c = v3(fmin(c.x, 8), fmin(c.y, 8), fmin(c.z, 8));
                L = add(L, c);
                nd = sample_lobe(n, 1, rng);
                T = vmul(T, s.alb);
                spec = 0; delta = 0;
            }
        }
        if (depth >= 3) {
            double q = fmin(0.95, maxc(T));
            if (rnd(rng) > q) break;
            T = mul(T, 1.0 / q);
        }
        o = po;
        d = nd;
    }
    return L;
}

/* ------------------------------------------------------------------ 场景搭建 */
static void build_scene(void) {
    /* 太阳：从窗外左后上方照进来 */
    SUN_DIR = norm(v3(-0.814, 0.5425, -0.2087));
    double ang = 0.55 * M_PI / 180.0;
    SUN_COS = cos(ang);
    SUN_E = mul(v3(1.0, 0.90, 0.76), 26.0);
    SUN_L = mul(SUN_E, 1.0 / (2 * M_PI * (1 - SUN_COS)));

    /* ---- 房间壳体：x∈[0,5], y∈[0,2.7], z∈[0,6]；左墙 x=0 开窗 ---- */
    begin_group();
    box(v3(-0.3, -0.2, -0.3), v3(5.3, 0.0, 6.3), M_FLOOR);
    box(v3(-0.3, 2.7, -0.3), v3(5.3, 2.9, 6.3), M_CEIL);
    box(v3(-0.3, 0.0, -0.2), v3(5.3, 2.7, 0.0), M_WALL);  /* 后墙 */
    box(v3(-0.3, 0.0, 6.0), v3(5.3, 2.7, 6.2), M_WALL);   /* 前墙（相机身后） */
    box(v3(5.0, 0.0, -0.3), v3(5.2, 2.7, 6.3), M_WALL);   /* 右墙 */
    box(v3(-WALL_T, 0.0, -0.3), v3(0.0, 2.7, WIN_Z0), M_WALL);
    box(v3(-WALL_T, 0.0, WIN_Z1), v3(0.0, 2.7, 6.3), M_WALL);
    box(v3(-WALL_T, 0.0, WIN_Z0), v3(0.0, WIN_Y0, WIN_Z1), M_WALL);
    box(v3(-WALL_T, WIN_Y1, WIN_Z0), v3(0.0, 2.7, WIN_Z1), M_WALL);
    /* 窗框与窗台 */
    {
        double x0 = -0.13, x1 = -0.06, f = 0.055;
        box(v3(x0, WIN_Y0, WIN_Z0), v3(x1, WIN_Y0 + f, WIN_Z1), M_TRIM);
        box(v3(x0, WIN_Y1 - f, WIN_Z0), v3(x1, WIN_Y1, WIN_Z1), M_TRIM);
        box(v3(x0, WIN_Y0, WIN_Z0), v3(x1, WIN_Y1, WIN_Z0 + f), M_TRIM);
        box(v3(x0, WIN_Y0, WIN_Z1 - f), v3(x1, WIN_Y1, WIN_Z1), M_TRIM);
        double zm = 0.5 * (WIN_Z0 + WIN_Z1);
        box(v3(x0, WIN_Y0, zm - 0.025), v3(x1, WIN_Y1, zm + 0.025), M_TRIM);
        box(v3(x0, 1.86, WIN_Z0), v3(x1, 1.905, WIN_Z1), M_TRIM);
        box(v3(-0.10, WIN_Y0 - 0.03, WIN_Z0 - 0.06), v3(0.05, WIN_Y0 + 0.012, WIN_Z1 + 0.06), M_TRIM);
    }
    /* 踢脚线 */
    box(v3(0.0, 0.0, 0.0), v3(5.0, 0.09, 0.014), M_TRIM);
    box(v3(0.0, 0.0, 0.0), v3(0.014, 0.09, 6.0), M_TRIM);
    box(v3(4.986, 0.0, 0.0), v3(5.0, 0.09, 6.0), M_TRIM);
    end_group();

    /* ---- 沙发（靠后墙） ---- */
    begin_group();
    {
        double cx = 3.15;
        for (int i = -1; i <= 1; i += 2)
            for (int j = 0; j < 2; j++)
                cyl(v3(cx + i * 1.0, 0.055, j ? 0.86 : 0.12), 0.022, 0.055, 0.004, M_WALNUT);
        rbox(v3(cx, 0.25, 0.49), v3(1.12, 0.14, 0.46), 0.035, M_SOFA, 0, 0, 0);
        rbox(v3(cx - 1.0, 0.43, 0.49), v3(0.13, 0.24, 0.46), 0.065, M_SOFA, 0, 0, 0);
        rbox(v3(cx + 1.0, 0.43, 0.49), v3(0.13, 0.24, 0.46), 0.065, M_SOFA, 0, 0, 0);
        rbox(v3(cx, 0.62, 0.15), v3(0.90, 0.26, 0.12), 0.06, M_SOFA, 0, 0, 0);
        rbox(v3(cx - 0.437, 0.455, 0.60), v3(0.43, 0.08, 0.36), 0.06, M_SOFA, 0, 0, 0);
        rbox(v3(cx + 0.437, 0.455, 0.60), v3(0.43, 0.08, 0.36), 0.06, M_SOFA, 0, 0, 0);
        rbox(v3(cx - 0.437, 0.74, 0.34), v3(0.42, 0.20, 0.085), 0.08, M_SOFA, 0, -0.20, 0);
        rbox(v3(cx + 0.437, 0.74, 0.34), v3(0.42, 0.20, 0.085), 0.08, M_SOFA, 0, -0.20, 0);
        rbox(v3(cx - 0.70, 0.70, 0.50), v3(0.20, 0.19, 0.06), 0.075, M_PILLOW_MUSTARD, 0.35, -0.25, 0.12);
        rbox(v3(cx + 0.72, 0.70, 0.50), v3(0.21, 0.20, 0.06), 0.075, M_PILLOW_CREAM, -0.30, -0.25, -0.10);
    }
    end_group();

    /* ---- 地毯 + 茶几 + 桌上摆件 ---- */
    begin_group();
    rbox(v3(3.0, 0.007, 1.9), v3(1.4, 0.007, 1.0), 0.006, M_RUG, 0, 0, 0);
    {
        double cx = 3.1, cz = 1.85, top = 0.40;
        rbox(v3(cx, top, cz), v3(0.60, 0.022, 0.32), 0.012, M_WALNUT, 0, 0, 0);
        rbox(v3(cx, 0.13, cz), v3(0.55, 0.012, 0.27), 0.006, M_WALNUT, 0, 0, 0);
        for (int i = -1; i <= 1; i += 2)
            for (int j = -1; j <= 1; j += 2)
                cyl(v3(cx + i * 0.53, 0.195, cz + j * 0.25), 0.02, 0.195, 0.004, M_WALNUT);
        double y = top + 0.022;
        Obj *b;
        b = rbox(v3(2.82, y + 0.018, 1.80), v3(0.15, 0.018, 0.11), 0.004, M_BOOK_FLAT, 0.14, 0, 0);
        b->col = v3(0.30, 0.06, 0.05);
        b = rbox(v3(2.83, y + 0.036 + 0.014, 1.81), v3(0.13, 0.014, 0.095), 0.004, M_BOOK_FLAT, -0.08, 0, 0);
        b->col = v3(0.62, 0.58, 0.48);
        b = rbox(v3(2.81, y + 0.064 + 0.011, 1.80), v3(0.11, 0.011, 0.08), 0.003, M_BOOK_FLAT, 0.30, 0, 0);
        b->col = v3(0.05, 0.10, 0.09);
        cup(v3(3.42, y + 0.11, 1.92), 0.062, 0.11, 0.007, M_CERAMIC);
        cyl(v3(3.62, y + 0.012, 1.76), 0.09, 0.012, 0.006, M_BLACK);
        cup(v3(3.62, y + 0.024 + 0.045, 1.76), 0.042, 0.045, 0.004, M_CERAMIC);
        /* 花瓶里的枝条 */
        V base = v3(3.42, y + 0.15, 1.92);
        cap(base, v3(3.36, y + 0.55, 1.88), 0.004, M_STEM);
        cap(base, v3(3.50, y + 0.50, 1.97), 0.004, M_STEM);
        cap(base, v3(3.45, y + 0.62, 1.86), 0.004, M_STEM);
    }
    end_group();

    /* ---- 落地灯 ---- */
    begin_group();
    {
        double lx = 1.72, lz = 0.38;
        cyl(v3(lx, 0.012, lz), 0.14, 0.012, 0.006, M_BLACK);
        cap(v3(lx, 0.02, lz), v3(lx, 1.42, lz), 0.011, M_BRASS);
        tube(v3(lx, 1.55, lz), 0.21, 0.15, 0.004, M_SHADE);
        cap(v3(lx - 0.2, 1.62, lz), v3(lx + 0.2, 1.62, lz), 0.004, M_BRASS);
    }
    end_group();

    /* ---- 盆栽（琴叶榕风格） ---- */
    begin_group();
    {
        double px = 0.52, pz = 0.36;
        cup(v3(px, 0.20, pz), 0.17, 0.20, 0.012, M_POT);
        cyl(v3(px, 0.34, pz), 0.155, 0.01, 0.004, M_SOIL);
        V t0 = v3(px, 0.34, pz), t1 = v3(px + 0.03, 0.95, pz + 0.02), t2 = v3(px - 0.02, 1.55, pz + 0.03);
        cap(t0, t1, 0.014, M_STEM);
        cap(t1, t2, 0.011, M_STEM);
        Rng r = {12345};
        pcg(&r);
        int N = 30;
        for (int i = 0; i < N; i++) {
            double t = pow((i + 0.5) / N, 0.7);
            double hgt = 0.60 + t * 0.95;
            V tp = hgt < 0.95 ? mixv(t0, t1, (hgt - 0.34) / (0.95 - 0.34)) : mixv(t1, t2, (hgt - 0.95) / 0.60);
            double az = i * 2.39996 + (rnd(&r) - 0.5) * 0.6;
            V out = v3(cos(az), 0, sin(az));
            double el = -0.25 + 0.75 * rnd(&r) + 0.2 * t;
            V f = norm(add(mul(out, cos(el)), v3(0, sin(el), 0)));
            double L = 0.10 + 0.05 * rnd(&r);
            V pet_end = add(tp, add(mul(out, 0.035), v3(0, 0.02, 0)));
            V c = add(pet_end, mul(f, L * 0.95));
            if (c.z < 0.10) c.z = 0.10;
            if (c.x < 0.20) c.x = 0.20;
            cap(tp, pet_end, 0.004, M_STEM);
            double g = rnd(&r);
            V col = mixv(v3(0.035, 0.10, 0.025), v3(0.09, 0.20, 0.05), g);
            leaf(c, f, (rnd(&r) - 0.5) * 1.0, v3(0.065 + 0.02 * rnd(&r), 0.0045, L), 1.6, col);
        }
    }
    end_group();

    /* ---- 窗帘 + 窗帘杆 ---- */
    begin_group();
    curtain(v3(0.10, 1.33, 0.30), 1.29, 0.21, 0.035, 2 * M_PI / 0.115, 0.004, 0.3, M_CURTAIN);
    curtain(v3(0.10, 1.33, 2.62), 1.29, 0.24, 0.035, 2 * M_PI / 0.12, 0.004, 1.7, M_CURTAIN);
    cap(v3(0.11, 2.64, 0.06), v3(0.11, 2.64, 2.94), 0.011, M_BLACK);
    end_group();

    /* ---- 墙上挂画 ---- */
    begin_group();
    {
        Obj *a = rbox(v3(2.82, 1.56, 0.016), v3(0.28, 0.36, 0.016), 0.003, M_ART, 0, 0, 0);
        a->id = 0;
        a = rbox(v3(3.48, 1.56, 0.016), v3(0.28, 0.36, 0.016), 0.003, M_ART, 0, 0, 0);
        a->id = 1;
    }
    end_group();

    /* ---- 书架（右后角） ---- */
    begin_group();
    {
        double x0 = 4.40, x1 = 4.98, d = 0.34, top = 1.92;
        box(v3(x0, 0, 0), v3(x0 + 0.022, top, d), M_OAK);
        box(v3(x1 - 0.022, 0, 0), v3(x1, top, d), M_OAK);
        box(v3(x0, top - 0.022, 0), v3(x1, top, d), M_OAK);
        box(v3(x0, 0, 0), v3(x1, 0.08, d), M_OAK);
        box(v3(x0, 0, 0), v3(x1, top, 0.012), M_OAK);
        double shelves[5] = {0.08, 0.44, 0.80, 1.16, 1.52};
        for (int s = 1; s < 5; s++) box(v3(x0, shelves[s], 0), v3(x1, shelves[s] + 0.022, d), M_OAK);
        V pal[10] = {v3(0.05, 0.07, 0.15), v3(0.30, 0.05, 0.05), v3(0.55, 0.38, 0.10), v3(0.62, 0.58, 0.48),
                     v3(0.06, 0.15, 0.08), v3(0.25, 0.25, 0.24), v3(0.03, 0.03, 0.03), v3(0.45, 0.17, 0.08),
                     v3(0.70, 0.68, 0.64), v3(0.20, 0.28, 0.35)};
        Rng r = {987654};
        pcg(&r);
        for (int s = 0; s < 5; s++) {
            double fy = shelves[s] + (s ? 0.022 : 0.0);
            double ceil_y = (s < 4 ? shelves[s + 1] : top - 0.022);
            double x = x0 + 0.026;
            double xend = x1 - 0.026;
            int gap_at = (int)(rnd(&r) * 3);
            int k = 0;
            while (x < xend - 0.02) {
                if (k == gap_at && s % 2 == 1) {
                    /* 留一个位置放小摆件 */
                    double rad = 0.045;
                    if (x + 2 * rad + 0.02 > xend) break;
                    cyl(v3(x + rad + 0.01, fy + 0.07, d - 0.12), rad, 0.07, 0.02, s == 1 ? M_CERAMIC : M_BLACK);
                    x += 2 * rad + 0.03;
                    k++;
                    continue;
                }
                double w = 0.018 + 0.026 * rnd(&r);
                double hgt = fmin(ceil_y - fy - 0.02, 0.17 + 0.13 * rnd(&r));
                double dep = 0.15 + 0.08 * rnd(&r);
                if (x + w > xend) break;
                double zf = d - 0.005 - 0.015 * rnd(&r);
                Obj *b = box(v3(x, fy, zf - dep), v3(x + w, fy + hgt, zf), M_BOOK_UP);
                b->col = mul(pal[(int)(rnd(&r) * 10) % 10], 0.8 + 0.4 * rnd(&r));
                x += w + 0.001 + (rnd(&r) < 0.1 ? 0.01 : 0);
                k++;
            }
        }
    }
    end_group();
}

/* ------------------------------------------------------------------ 主程序 */
int main(int argc, char **argv) {
    if (argc < 5) {
        fprintf(stderr, "usage: %s W H SPP out.bin\n", argv[0]);
        return 1;
    }
    int W = atoi(argv[1]), H = atoi(argv[2]), SPP = atoi(argv[3]);
    const char *out = argv[4];
    build_scene();
    fprintf(stderr, "objects: %d, groups: %d\n", nobj, ngroup);

    V cam = v3(3.9, 1.25, 5.6);
    V fwd = norm(v3(-0.36, 0, -1));
    V right = norm(cross(fwd, v3(0, 1, 0)));
    V up = cross(right, fwd);
    double tanH = tan(34.0 * M_PI / 180.0);
    double tanV = tanH * H / W;
    double shift = -0.12; /* 移轴：保持竖线垂直，同时多拍一些地面 */

    float *col = calloc((size_t)W * H * 3, sizeof(float));
    float *alb = calloc((size_t)W * H * 3, sizeof(float));
    float *nrm = calloc((size_t)W * H * 3, sizeof(float));
    float *dep = calloc((size_t)W * H, sizeof(float));
    int done = 0;
    double t_start = omp_get_wtime();

#pragma omp parallel for schedule(dynamic, 1)
    for (int y = 0; y < H; y++) {
        for (int x = 0; x < W; x++) {
            Rng rng;
            rng.s = ((uint64_t)(y * W + x) + 1) * 0x9E3779B97F4A7C15ULL ^ 0x5DEECE66DULL;
            pcg(&rng); pcg(&rng);
            V sum = v3(0, 0, 0), sa = v3(0, 0, 0), sn = v3(0, 0, 0);
            double sd = 0;
            for (int s = 0; s < SPP; s++) {
                double u1 = 2 * rnd(&rng), u2 = 2 * rnd(&rng);
                double dx = u1 < 1 ? sqrt(u1) - 1 : 1 - sqrt(2 - u1);
                double dy = u2 < 1 ? sqrt(u2) - 1 : 1 - sqrt(2 - u2);
                double sx = ((x + 0.5 + 0.7 * dx) / W) * 2 - 1;
                double sy = 1 - ((y + 0.5 + 0.7 * dy) / H) * 2;
                V dir = norm(add(fwd, add(mul(right, sx * tanH), mul(up, sy * tanV + shift))));
                Aov a = {{0, 0, 0}, {0, 0, 0}, 0, 0};
                V c = radiance(cam, dir, &rng, &a);
                c = v3(fmin(c.x, 40), fmin(c.y, 40), fmin(c.z, 40));
                sum = add(sum, c);
                sa = add(sa, a.alb);
                sn = add(sn, a.nrm);
                sd += a.depth;
            }
            size_t i = (size_t)y * W + x;
            double inv = 1.0 / SPP;
            col[3 * i] = sum.x * inv; col[3 * i + 1] = sum.y * inv; col[3 * i + 2] = sum.z * inv;
            alb[3 * i] = sa.x * inv; alb[3 * i + 1] = sa.y * inv; alb[3 * i + 2] = sa.z * inv;
            nrm[3 * i] = sn.x * inv; nrm[3 * i + 1] = sn.y * inv; nrm[3 * i + 2] = sn.z * inv;
            dep[i] = sd * inv;
        }
#pragma omp atomic
        done++;
        if (done % 40 == 0) {
            double el = omp_get_wtime() - t_start;
            fprintf(stderr, "\r%d/%d rows, %.0fs elapsed, ~%.0fs left   ", done, H, el, el / done * (H - done));
        }
    }
    fprintf(stderr, "\nrender time: %.1fs\n", omp_get_wtime() - t_start);

    FILE *f = fopen(out, "wb");
    int32_t wh[2] = {W, H};
    fwrite(wh, sizeof(int32_t), 2, f);
    fwrite(col, sizeof(float), (size_t)W * H * 3, f);
    fwrite(alb, sizeof(float), (size_t)W * H * 3, f);
    fwrite(nrm, sizeof(float), (size_t)W * H * 3, f);
    fwrite(dep, sizeof(float), (size_t)W * H, f);
    fclose(f);
    return 0;
}
