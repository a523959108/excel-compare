import pandas as pd

class CompareLogic:
    @staticmethod
    def check_duplicate_pk(df, pk_cols):
        if df is None or df.empty or not pk_cols:
            return []
        for col in pk_cols:
            if col not in df.columns:
                return []
        duplicates = df[df.duplicated(subset=pk_cols, keep=False)]
        if len(duplicates) == 0:
            return []
        return duplicates[pk_cols].drop_duplicates().values.tolist()
    
    @staticmethod
    def compare_dfs(left_df, right_df, pk_cols, sub_pk_cols=None, compare_cols=None, output_cols=None, compare_by_row=False):
        if left_df is None or left_df.empty:
            raise Exception("左表不能为空，请检查Sheet内容是否为空")
        if right_df is None:
            right_df = pd.DataFrame()
        if sub_pk_cols is None:
            sub_pk_cols = []
        
        left_df = left_df.copy()
        right_df = right_df.copy()
        
        # 输出列完全和左表一致，顺序不变
        output_cols = left_df.columns.tolist()
        if not output_cols:
            raise Exception("左表没有有效列")
        
        compare_cols = [col for col in compare_cols if col in output_cols and (col in right_df.columns or right_df.empty)]
        
        if not compare_cols and not right_df.empty:
            raise Exception("两个表没有共同的可比对字段，请检查列名是否一致")
        
        if not compare_by_row:
            if not pk_cols:
                raise Exception("请至少选择一个主匹配主键")
            for col in pk_cols:
                if col not in left_df.columns:
                    raise Exception(f"主主键列 {col} 不存在于左表")
                if col not in right_df.columns and not right_df.empty:
                    raise Exception(f"主主键列 {col} 不存在于右表")
            for col in sub_pk_cols:
                if col not in left_df.columns:
                    raise Exception(f"辅助主键列 {col} 不存在于左表")
                if col not in right_df.columns and not right_df.empty:
                    raise Exception(f"辅助主键列 {col} 不存在于右表")
        
        # 合并主主键和辅助主键作为完整匹配键，去重，顺序和左表完全一致
        full_pk_cols = list(dict.fromkeys([col for col in left_df.columns if col in (pk_cols + sub_pk_cols)]))
        
        diff_columns = compare_cols
        result = []
        
        if compare_by_row:
            max_rows = max(len(left_df), len(right_df))
            for i in range(len(left_df)):
                left_row = left_df.iloc[i]
                row_data = {col: str(left_row[col]) for col in output_cols}
                row_status = 'unchanged'
                row_data['__diff_info'] = {}
                
                if i < len(right_df):
                    right_row = right_df.iloc[i]
                    for col in compare_cols:
                        left_val = str(left_row[col])
                        right_val = str(right_row.get(col, ''))
                        if left_val != right_val:
                            row_data[col] = f"{left_val} (修改: {right_val})"
                            row_data['__diff_info'][col] = (left_val, right_val)
                            row_status = 'modified'
                else:
                    row_status = 'deleted'
                
                row_data['__status'] = row_status
                result.append(row_data)
            
            # 追加右表多出的行
            for i in range(len(left_df), len(right_df)):
                right_row = right_df.iloc[i]
                row_data = {}
                for col in output_cols:
                    row_data[col] = str(right_row.get(col, ''))
                row_data['__status'] = 'added'
                row_data['__diff_info'] = {}
                result.append(row_data)
        else:
            # 按主键比对，完全保留左表顺序和结构
            right_groups = {}
            right_full_pk_set = set()
            if not right_df.empty:
                # 右表主键顺序强制和左表一致
                right_groups = right_df.groupby(full_pk_cols)
                right_full_pk_set = set(right_groups.groups.keys())
            
            # 第一步：全部保留左表的行，顺序不变
            for _, left_row in left_df.iterrows():
                # 主键值统一转字符串，避免类型不匹配
                full_pk_tuple = tuple(str(left_row[col]).strip() for col in full_pk_cols)
                row_data = {col: str(left_row[col]) for col in output_cols}
                row_status = 'unchanged'
                row_data['__diff_info'] = {}
                
                if full_pk_tuple in right_full_pk_set:
                    right_row = right_groups.get_group(full_pk_tuple).iloc[0]
                    for col in compare_cols:
                        left_val = str(left_row[col])
                        right_val = str(right_row[col])
                        if left_val != right_val:
                            row_data[col] = f"{left_val} (修改: {right_val})"
                            row_data['__diff_info'][col] = (left_val, right_val)
                            row_status = 'modified'
                else:
                    # 完整主键匹配不到，标删除
                    row_status = 'deleted'
                
                row_data['__status'] = row_status
                result.append(row_data)
            
            # 第二步：追加右表有左表没有的行
            left_full_pk_set = set(tuple(str(row[col]).strip() for col in full_pk_cols) for _, row in left_df.iterrows())
            for full_pk_tuple in right_full_pk_set:
                if full_pk_tuple not in left_full_pk_set:
                    right_rows = right_groups.get_group(full_pk_tuple)
                    for _, right_row in right_rows.iterrows():
                        row_data = {}
                        for col in output_cols:
                            row_data[col] = str(right_row.get(col, ''))
                        row_data['__status'] = 'added'
                        row_data['__diff_info'] = {}
                        result.append(row_data)
        
        result_df = pd.DataFrame(result)
        return result_df, diff_columns