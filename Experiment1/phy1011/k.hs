-- 运行方法：ghci k --

avg :: [Float] -> Float
avg xs = sum xs / fromIntegral (length xs)

t_0s = [7.948, 8.892, 7.932, 7.220, 7.940]
t_1s = [15.380, 15.512, 15.514, 15.520, 15.000]
t_2s = [16.375, 16.369, 16.372, 16.134, 16.207]

t_0 = 1 / 10 * avg t_0s
t_1 = 1 / 10 * avg t_1s
t_2 = 1 / 10 * avg t_2s

m_1 = 1.11139
d_1 = 0.10026
m_2 = 0.70463
d_2inner = 0.09406
d_2outer = 0.10026
i_1 = 1/8 * m_1 * d_1 ** 2

k = 4 * pi ** 2 * i_1 / (t_1 ** 2 - t_0 ** 2)

i_2 = k / (4 * pi ** 2) * (t_2 ** 2 - t_0 **2)
i_2_theory = 1 / 8 * m_2 * (d_2inner ** 2 + d_2outer ** 2)

relativeError measure theory = abs (measure - theory) / theory
-- toPercent x = (show x) ++ "%"
-- relativeErrorPercent = toPercent . relativeError

-- 相对误差
epsilon_2 = relativeError i_2 i_2_theory

-- 思考题

l = 36.80 --cm
delta_l = 0.3 --cm
u_l = delta_l / sqrt 3
ru_l = u_l / l

h = 103.90
delta_h = 0.5
u_h = delta_h / sqrt 3
ru_h = u_h / h

b = 2.00
delta_b = 0.02
u_b = delta_b / sqrt 3
ru_b = u_b / b

ru_d = 3.903844 / 0.782 / 1000
ru_c = 2.940193 / 0.779 / 1000
