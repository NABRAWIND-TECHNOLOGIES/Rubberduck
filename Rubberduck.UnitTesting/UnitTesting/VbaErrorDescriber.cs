using System.Runtime.InteropServices;

namespace Rubberduck.UnitTesting
{
    /// <summary>
    /// Pure helper (D12, headless-test-runner) that turns a <see cref="COMException"/> raised
    /// while running VBA code into a diagnostic message. When the HRESULT's facility is
    /// FACILITY_CONTROL (0x800A) -- the facility VBA uses to re-surface its own runtime error
    /// numbers across the COM boundary -- the low 16 bits of the HRESULT are the VBA error
    /// number, decoded and paired with the exception's description. Any other HRESULT still
    /// surfaces its hex value and description instead of a fixed generic string, since the
    /// exception itself is always available even when it cannot be decoded as a VBA error.
    /// </summary>
    internal static class VbaErrorDescriber
    {
        private const int HResultFacilityMask = unchecked((int)0xFFFF0000);
        private const int VbaFacilityControl = unchecked((int)0x800A0000);

        public static string Describe(COMException exception)
        {
            var hresult = exception.ErrorCode;

            if ((hresult & HResultFacilityMask) == VbaFacilityControl)
            {
                var vbaErrorNumber = hresult & 0xFFFF;
                return $"VBA runtime error {vbaErrorNumber}: {exception.Message}";
            }

            return $"COM error 0x{unchecked((uint)hresult):X8}: {exception.Message}";
        }
    }
}
