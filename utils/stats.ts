import { supabase } from "@/lib/supabase";

export const logConversion = async (fromFormat: string, toFormat: string) => {
  await supabase.from("conversions").insert([
    {
      from_format: fromFormat,
      to_format: toFormat,
      created_at: new Date().toISOString(),
    },
  ]);
};