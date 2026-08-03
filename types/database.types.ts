export type Json =
  | string
  | number
  | boolean
  | null
  | { [key: string]: Json | undefined }
  | Json[]

export type Database = {
  // Allows to automatically instantiate createClient with right options
  // instead of createClient<Database, { PostgrestVersion: 'XX' }>(URL, KEY)
  __InternalSupabase: {
    PostgrestVersion: "14.5"
  }
  public: {
    Tables: {
      carrier: {
        Row: {
          carrier_code: string
          carrier_name: string
        }
        Insert: {
          carrier_code: string
          carrier_name: string
        }
        Update: {
          carrier_code?: string
          carrier_name?: string
        }
        Relationships: []
      }
      location: {
        Row: {
          location_id: number
          location_key: string
          pickup_address: string | null
          pickup_address_state: string | null
          pickup_location: string
          status: string | null
        }
        Insert: {
          location_id?: number
          location_key: string
          pickup_address?: string | null
          pickup_address_state?: string | null
          pickup_location: string
          status?: string | null
        }
        Update: {
          location_id?: number
          location_key?: string
          pickup_address?: string | null
          pickup_address_state?: string | null
          pickup_location?: string
          status?: string | null
        }
        Relationships: []
      }
      partner: {
        Row: {
          location_key: string | null
          partner_id: string
          partner_name: string
        }
        Insert: {
          location_key?: string | null
          partner_id: string
          partner_name: string
        }
        Update: {
          location_key?: string | null
          partner_id?: string
          partner_name?: string
        }
        Relationships: []
      }
      pickup_customs_broker: {
        Row: {
          bol_picture_link: string | null
          id: number
          inbound_time: string | null
          main_air_waybill: string | null
          pickup_date: string | null
          pickup_location: string | null
          pickup_quantity: number | null
          quantity_unit: string | null
        }
        Insert: {
          bol_picture_link?: string | null
          id?: number
          inbound_time?: string | null
          main_air_waybill?: string | null
          pickup_date?: string | null
          pickup_location?: string | null
          pickup_quantity?: number | null
          quantity_unit?: string | null
        }
        Update: {
          bol_picture_link?: string | null
          id?: number
          inbound_time?: string | null
          main_air_waybill?: string | null
          pickup_date?: string | null
          pickup_location?: string | null
          pickup_quantity?: number | null
          quantity_unit?: string | null
        }
        Relationships: []
      }
      pickup_final_mile: {
        Row: {
          id: number
          location_key: string | null
          pickup_date: string | null
          pickup_quantity: number | null
          quantity_unit: string | null
        }
        Insert: {
          id?: number
          location_key?: string | null
          pickup_date?: string | null
          pickup_quantity?: number | null
          quantity_unit?: string | null
        }
        Update: {
          id?: number
          location_key?: string | null
          pickup_date?: string | null
          pickup_quantity?: number | null
          quantity_unit?: string | null
        }
        Relationships: []
      }
      pickup_gofo_cbt: {
        Row: {
          forecast_pickup_quantity: number | null
          id: number
          inbound_time: string | null
          pickup_date: string | null
          pickup_quantity: number | null
          quantity_unit: string | null
          scan_completion_time: string | null
        }
        Insert: {
          forecast_pickup_quantity?: number | null
          id?: number
          inbound_time?: string | null
          pickup_date?: string | null
          pickup_quantity?: number | null
          quantity_unit?: string | null
          scan_completion_time?: string | null
        }
        Update: {
          forecast_pickup_quantity?: number | null
          id?: number
          inbound_time?: string | null
          pickup_date?: string | null
          pickup_quantity?: number | null
          quantity_unit?: string | null
          scan_completion_time?: string | null
        }
        Relationships: []
      }
      pickup_local_merchant: {
        Row: {
          carrier_code: string | null
          id: number
          inbound_time: string | null
          location_key: string | null
          pickup_address: string | null
          pickup_address_state: string | null
          pickup_date: string | null
          pickup_location: string | null
          pickup_price: number | null
          pickup_quantity: number | null
          quantity_unit: string | null
          truck_type: string | null
        }
        Insert: {
          carrier_code?: string | null
          id?: number
          inbound_time?: string | null
          location_key?: string | null
          pickup_address?: string | null
          pickup_address_state?: string | null
          pickup_date?: string | null
          pickup_location?: string | null
          pickup_price?: number | null
          pickup_quantity?: number | null
          quantity_unit?: string | null
          truck_type?: string | null
        }
        Update: {
          carrier_code?: string | null
          id?: number
          inbound_time?: string | null
          location_key?: string | null
          pickup_address?: string | null
          pickup_address_state?: string | null
          pickup_date?: string | null
          pickup_location?: string | null
          pickup_price?: number | null
          pickup_quantity?: number | null
          quantity_unit?: string | null
          truck_type?: string | null
        }
        Relationships: []
      }
      pricing: {
        Row: {
          location_key: string
          pickup_carrier: string
          pickup_price: number
          pricing_id: number
          record_effective_date: string
          truck_type: string | null
        }
        Insert: {
          location_key: string
          pickup_carrier: string
          pickup_price: number
          pricing_id?: number
          record_effective_date: string
          truck_type?: string | null
        }
        Update: {
          location_key?: string
          pickup_carrier?: string
          pickup_price?: number
          pricing_id?: number
          record_effective_date?: string
          truck_type?: string | null
        }
        Relationships: [
          {
            foreignKeyName: "fk_pricing_carrier"
            columns: ["pickup_carrier"]
            isOneToOne: false
            referencedRelation: "carrier"
            referencedColumns: ["carrier_code"]
          },
          {
            foreignKeyName: "fk_pricing_location"
            columns: ["location_key"]
            isOneToOne: false
            referencedRelation: "location"
            referencedColumns: ["location_key"]
          },
        ]
      }
      pricing_group: {
        Row: {
          group_code: string
          id: number
          record_effective_date: string
        }
        Insert: {
          group_code: string
          id?: number
          record_effective_date: string
        }
        Update: {
          group_code?: string
          id?: number
          record_effective_date?: string
        }
        Relationships: []
      }
      pricing_group_location: {
        Row: {
          group_code: string
          location_key: string
          pricing_group_id: number
        }
        Insert: {
          group_code: string
          location_key: string
          pricing_group_id: number
        }
        Update: {
          group_code?: string
          location_key?: string
          pricing_group_id?: number
        }
        Relationships: []
      }
      pricing_group_rate: {
        Row: {
          pickup_price: number
          pricing_group_id: number
          record_effective_date: string
        }
        Insert: {
          pickup_price: number
          pricing_group_id: number
          record_effective_date: string
        }
        Update: {
          pickup_price?: number
          pricing_group_id?: number
          record_effective_date?: string
        }
        Relationships: [
          {
            foreignKeyName: "fk_pgr_group"
            columns: ["pricing_group_id"]
            isOneToOne: false
            referencedRelation: "pricing_group"
            referencedColumns: ["id"]
          },
        ]
      }
      rules: {
        Row: {
          carrier_code: string | null
          id: number
          location_key: string | null
          priority: number
          record_effective_date: string
          rule_type: string
          truck_rule: string | null
          use_db_price: boolean | null
        }
        Insert: {
          carrier_code?: string | null
          id?: number
          location_key?: string | null
          priority: number
          record_effective_date?: string
          rule_type: string
          truck_rule?: string | null
          use_db_price?: boolean | null
        }
        Update: {
          carrier_code?: string | null
          id?: number
          location_key?: string | null
          priority?: number
          record_effective_date?: string
          rule_type?: string
          truck_rule?: string | null
          use_db_price?: boolean | null
        }
        Relationships: [
          {
            foreignKeyName: "fk_rules_carrier"
            columns: ["carrier_code"]
            isOneToOne: false
            referencedRelation: "carrier"
            referencedColumns: ["carrier_code"]
          },
          {
            foreignKeyName: "fk_rules_location"
            columns: ["location_key"]
            isOneToOne: false
            referencedRelation: "location"
            referencedColumns: ["location_key"]
          },
        ]
      }
      schedule: {
        Row: {
          carrier_code: string
          inbound_warehouse_code: string | null
          location_key: string
          pickup_note: string | null
          pickup_schedule: string | null
          record_effective_date: string
          schedule_id: number
        }
        Insert: {
          carrier_code: string
          inbound_warehouse_code?: string | null
          location_key: string
          pickup_note?: string | null
          pickup_schedule?: string | null
          record_effective_date: string
          schedule_id?: number
        }
        Update: {
          carrier_code?: string
          inbound_warehouse_code?: string | null
          location_key?: string
          pickup_note?: string | null
          pickup_schedule?: string | null
          record_effective_date?: string
          schedule_id?: number
        }
        Relationships: [
          {
            foreignKeyName: "fk_schedule_carrier"
            columns: ["carrier_code"]
            isOneToOne: false
            referencedRelation: "carrier"
            referencedColumns: ["carrier_code"]
          },
          {
            foreignKeyName: "fk_schedule_location"
            columns: ["location_key"]
            isOneToOne: false
            referencedRelation: "location"
            referencedColumns: ["location_key"]
          },
        ]
      }
    }
    Views: {
      final_mile_view: {
        Row: {
          daily_total_price: number | null
          location_key: string | null
          pickup_date: string | null
          pickup_quantity: number | null
          price_display: string | null
        }
        Relationships: []
      }
    }
    Functions: {
      pickup_daily_cost_summary: {
        Args: {
          input_date: string
        }
        Returns: {
          amount: number
          charge_type: string
          group_code: string
          pickup_count: number
        }[]
      }
      pickup_group_charges_for_date: {
        Args: {
          input_date: string
        }
        Returns: {
          bundle_price: number
          group_code: string
          member_locations: string[]
          pickup_count: number
          pricing_group_id: number
        }[]
      }
      pickup_lookup: {
        Args: {
          input_date: string
          input_pickup_location: string
          input_truck_type: string
        }
        Returns: {
          carrier_code: string
          location_key: string
          pickup_address: string
          pickup_address_state: string
          pickup_location: string
          pickup_price: number
          truck_type_final: string
        }[]
      }
    }
    Enums: {
      [_ in never]: never
    }
    CompositeTypes: {
      [_ in never]: never
    }
  }
}

type DatabaseWithoutInternals = Omit<Database, "__InternalSupabase">

type DefaultSchema = DatabaseWithoutInternals[Extract<keyof Database, "public">]

export type Tables<
  DefaultSchemaTableNameOrOptions extends
    | keyof (DefaultSchema["Tables"] & DefaultSchema["Views"])
    | { schema: keyof DatabaseWithoutInternals },
  TableName extends DefaultSchemaTableNameOrOptions extends {
    schema: keyof DatabaseWithoutInternals
  }
    ? keyof (DatabaseWithoutInternals[DefaultSchemaTableNameOrOptions["schema"]]["Tables"] &
        DatabaseWithoutInternals[DefaultSchemaTableNameOrOptions["schema"]]["Views"])
    : never = never,
> = DefaultSchemaTableNameOrOptions extends {
  schema: keyof DatabaseWithoutInternals
}
  ? (DatabaseWithoutInternals[DefaultSchemaTableNameOrOptions["schema"]]["Tables"] &
      DatabaseWithoutInternals[DefaultSchemaTableNameOrOptions["schema"]]["Views"])[TableName] extends {
      Row: infer R
    }
    ? R
    : never
  : DefaultSchemaTableNameOrOptions extends keyof (DefaultSchema["Tables"] &
        DefaultSchema["Views"])
    ? (DefaultSchema["Tables"] &
        DefaultSchema["Views"])[DefaultSchemaTableNameOrOptions] extends {
        Row: infer R
      }
      ? R
      : never
    : never

export type TablesInsert<
  DefaultSchemaTableNameOrOptions extends
    | keyof DefaultSchema["Tables"]
    | { schema: keyof DatabaseWithoutInternals },
  TableName extends DefaultSchemaTableNameOrOptions extends {
    schema: keyof DatabaseWithoutInternals
  }
    ? keyof DatabaseWithoutInternals[DefaultSchemaTableNameOrOptions["schema"]]["Tables"]
    : never = never,
> = DefaultSchemaTableNameOrOptions extends {
  schema: keyof DatabaseWithoutInternals
}
  ? DatabaseWithoutInternals[DefaultSchemaTableNameOrOptions["schema"]]["Tables"][TableName] extends {
      Insert: infer I
    }
    ? I
    : never
  : DefaultSchemaTableNameOrOptions extends keyof DefaultSchema["Tables"]
    ? DefaultSchema["Tables"][DefaultSchemaTableNameOrOptions] extends {
        Insert: infer I
      }
      ? I
      : never
    : never

export type TablesUpdate<
  DefaultSchemaTableNameOrOptions extends
    | keyof DefaultSchema["Tables"]
    | { schema: keyof DatabaseWithoutInternals },
  TableName extends DefaultSchemaTableNameOrOptions extends {
    schema: keyof DatabaseWithoutInternals
  }
    ? keyof DatabaseWithoutInternals[DefaultSchemaTableNameOrOptions["schema"]]["Tables"]
    : never = never,
> = DefaultSchemaTableNameOrOptions extends {
  schema: keyof DatabaseWithoutInternals
}
  ? DatabaseWithoutInternals[DefaultSchemaTableNameOrOptions["schema"]]["Tables"][TableName] extends {
      Update: infer U
    }
    ? U
    : never
  : DefaultSchemaTableNameOrOptions extends keyof DefaultSchema["Tables"]
    ? DefaultSchema["Tables"][DefaultSchemaTableNameOrOptions] extends {
        Update: infer U
      }
      ? U
      : never
    : never

export type Enums<
  DefaultSchemaEnumNameOrOptions extends
    | keyof DefaultSchema["Enums"]
    | { schema: keyof DatabaseWithoutInternals },
  EnumName extends DefaultSchemaEnumNameOrOptions extends {
    schema: keyof DatabaseWithoutInternals
  }
    ? keyof DatabaseWithoutInternals[DefaultSchemaEnumNameOrOptions["schema"]]["Enums"]
    : never = never,
> = DefaultSchemaEnumNameOrOptions extends {
  schema: keyof DatabaseWithoutInternals
}
  ? DatabaseWithoutInternals[DefaultSchemaEnumNameOrOptions["schema"]]["Enums"][EnumName]
  : DefaultSchemaEnumNameOrOptions extends keyof DefaultSchema["Enums"]
    ? DefaultSchema["Enums"][DefaultSchemaEnumNameOrOptions]
    : never

export type CompositeTypes<
  PublicCompositeTypeNameOrOptions extends
    | keyof DefaultSchema["CompositeTypes"]
    | { schema: keyof DatabaseWithoutInternals },
  CompositeTypeName extends PublicCompositeTypeNameOrOptions extends {
    schema: keyof DatabaseWithoutInternals
  }
    ? keyof DatabaseWithoutInternals[PublicCompositeTypeNameOrOptions["schema"]]["CompositeTypes"]
    : never = never,
> = PublicCompositeTypeNameOrOptions extends {
  schema: keyof DatabaseWithoutInternals
}
  ? DatabaseWithoutInternals[PublicCompositeTypeNameOrOptions["schema"]]["CompositeTypes"][CompositeTypeName]
  : PublicCompositeTypeNameOrOptions extends keyof DefaultSchema["CompositeTypes"]
    ? DefaultSchema["CompositeTypes"][PublicCompositeTypeNameOrOptions]
    : never

export const Constants = {
  public: {
    Enums: {},
  },
} as const
