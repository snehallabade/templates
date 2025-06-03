import { createClient } from '@supabase/supabase-js';

const supabaseUrl = 'https://werpkoihxashwuqrukda.supabase.co';
const supabaseKey = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6IndlcnBrb2loeGFzaHd1cXJ1a2RhIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NDgyNjQ1MDEsImV4cCI6MjA2Mzg0MDUwMX0.VT-b5guX7E2rIzbTeXZ-NiI2PpdHPPcNFu5FLo1YbzU';
export const supabase = createClient(supabaseUrl, supabaseKey);
