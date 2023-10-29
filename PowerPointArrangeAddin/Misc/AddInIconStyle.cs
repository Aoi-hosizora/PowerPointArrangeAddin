namespace PowerPointArrangeAddin.Misc {

    public enum AddInIconStyle {
        Office2013,
        Office2010
    }

    public static class AddInIconStyleExtension {

        public static AddInIconStyle ToAddInIconStyle(this string iconStyle) {
            return iconStyle switch {
                "2013" => AddInIconStyle.Office2013,
                "2010" => AddInIconStyle.Office2010,
                _ => AddInIconStyle.Office2013
            };
        }

        public static string ToIconStyleString(this AddInIconStyle iconStyle) {
            return iconStyle switch {
                AddInIconStyle.Office2013 => "2013",
                AddInIconStyle.Office2010 => "2010",
                _ => "2013"
            };
        }

        public static AddInIconStyle ToAddInIconStyle(this int index) {
            return index switch {
                0 => AddInIconStyle.Office2013,
                1 => AddInIconStyle.Office2010,
                _ => AddInIconStyle.Office2013
            };
        }

        public static int ToIconStyleIndex(this AddInIconStyle iconStyle) {
            return iconStyle switch {
                AddInIconStyle.Office2013 => 0,
                AddInIconStyle.Office2010 => 1,
                _ => 0
            };
        }

    }

}
