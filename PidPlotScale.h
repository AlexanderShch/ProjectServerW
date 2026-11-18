#pragma once

namespace ProjectServerW {
	constexpr int PidGridSubdivisions = 4;
	// Шаг сетки совпадает со смещением соседних кривых и оставляет
	// между полосами зазор хотя бы в один эквивалент ТЭНа.
	inline double PidPlotSpacing(double minimum, double maximum)
	{
		return System::Math::Max(4.0, System::Math::Ceiling(maximum - minimum) + 1.0);
	}
}
