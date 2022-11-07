#![allow(dead_code)]
use darling::FromField;
use proc_macro2::{Ident, TokenStream};
use quote::quote;
use syn::{
    Data, DataStruct, DeriveInput, Field, Fields, FieldsNamed, GenericArgument, Path, Type,
    TypePath,
};

pub mod macros {
    pub trait ExcelWriter {
        type Item;
        fn write_excel(list: Vec<Self::Item>, path: &std::path::Path);
    }

    pub trait ExcelReader {
        type Item;
        fn read(path: &std::path::Path) -> Vec<Self::Item>;

        fn check(path: &std::path::Path) -> bool {
            path.is_file()
        }

        fn read_file(path: &std::path::Path) -> Vec<Self::Item> {
            Self::check(path);
            Self::read(path)
        }

        fn default() {
            println!("ExcelReader default")
        }
    }
}

/// 用于捕获每个字段的 attributes 的结构
#[derive(Debug, Default, FromField)]
#[darling(default, attributes(excel))]
struct Opts {
    title: Option<String>,
    index: Option<u32>,
}

/// 我们需要的描述一个字段的所有信息
#[derive(Debug)]
struct Fd {
    name: Ident,
    ty: Type,
    is_option: bool,
    is_vec: bool,
    opts: Opts,
}

/// 把一个 Field 转换成 Fd
impl From<Field> for Fd {
    fn from(f: Field) -> Self {
        let (is_vec, _ty) = get_vec_inner(&f.ty);
        let (is_option, ty) = get_option_inner(&f.ty);
        // 从 Field 中读取 attributes 生成 Opts，如果没有使用缺省值
        let opts = Opts::from_field(&f).unwrap_or_default();
        Self {
            opts,
            // 此时，我们拿到的是 NamedFields，所以 ident 必然存在
            name: f.ident.unwrap(),
            is_option,
            is_vec,
            ty: ty.to_owned(),
        }
    }
}

/// 我们需要的描述一个 struct 的所有信息
#[derive(Debug)]
pub struct EasyExcelContext {
    name: Ident,
    fields: Vec<Fd>,
}

/// 把 DeriveInput 转换成 BuilderContext
impl From<DeriveInput> for EasyExcelContext {
    fn from(input: DeriveInput) -> Self {
        let name = input.ident;

        let fields = if let Data::Struct(DataStruct {
            fields: Fields::Named(FieldsNamed { named, .. }),
            ..
        }) = input.data
        {
            named
        } else {
            panic!("Unsupported data type");
        };

        let fds = fields.into_iter().map(Fd::from).collect();
        Self { name, fields: fds }
    }
}

impl EasyExcelContext {
    pub fn render(&self) -> syn::Result<TokenStream> {
        println!("{:#?}", self);
        let struct_name = &self.name;
        let header = &self.generate_header();
        let header_size = header.len();
        let fds = &self.fields;
        // let fields = &self.fields;
        let mut stream = TokenStream::new();
        for fd in fds.iter() {
            if !fd.is_vec {
                if let Some(title) = &fd.opts.title {
                    let fd_name = &fd.name;
                    stream.extend(quote!(
                        map.insert(#title.to_string(),temp.#fd_name.clone().to_string());
                    ));
                }
            }
        }
        let res = quote!(
            impl #struct_name {
                fn write_excel(list: Vec<#struct_name>, path: &std::path::Path) {
                    use umya_spreadsheet as xl_tool;
                    use std::collections::HashMap;
                    let mut book = xl_tool::new_file_empty_worksheet();
                    let sheet = book.new_sheet("Sheet1").unwrap();
                    let header = vec![
                        #(#header,)*
                    ];
                    for col in 1..=#header_size {
                        let cell = sheet.get_cell_by_column_and_row_mut(&(col as u32), &(1 as u32));
                        cell.set_value_from_string(header[col - 1]);
                    }
                    for row in 2..=list.len()+1 {
                        let temp = &list[row - 2];
                        let mut map: HashMap<String,String> = HashMap::new();
                        #stream
                        for col in 1..=#header_size {
                            let cell = sheet.get_cell_by_column_and_row_mut(&(col as u32), &(row as u32));
                            cell.set_value_from_string(map.get(header[col - 1]).unwrap());
                        }
                    }
                    let _ = xl_tool::writer::xlsx::write(&book, path);
                }
            }
        );
        Ok(res)
    }

    fn generate_header(&self) -> Vec<String> {
        // let mut header = Vec::new();
        let header: Vec<_> = self
            .fields
            .iter()
            .filter(|fd| !fd.is_vec)
            .filter_map(|fd| fd.opts.title.clone())
            .collect();
        header
        // quote!()
    }
}

// 如果是 T = Option<Inner>，返回 (true, Inner)；否则返回 (false, T)
fn get_option_inner(ty: &Type) -> (bool, &Type) {
    get_type_inner(ty, "Option")
}

// 如果是 T = Vec<Inner>，返回 (true, Inner)；否则返回 (false, T)
fn get_vec_inner(ty: &Type) -> (bool, &Type) {
    get_type_inner(ty, "Vec")
}

fn get_type_inner<'a>(ty: &'a Type, name: &str) -> (bool, &'a Type) {
    // 首先模式匹配出 segments
    if let Type::Path(TypePath {
        path: Path { segments, .. },
        ..
    }) = ty
    {
        if let Some(v) = segments.iter().next() {
            if v.ident == name {
                // 如果 PathSegment 第一个是 Option/Vec 等类型，那么它内部应该是 AngleBracketed，比如 <T>
                // 获取其第一个值，如果是 GenericArgument::Type，则返回
                let t = match &v.arguments {
                    syn::PathArguments::AngleBracketed(a) => match a.args.iter().next() {
                        Some(GenericArgument::Type(t)) => t,
                        _ => panic!("Not sure what to do with other GenericArgument"),
                    },
                    _ => panic!("Not sure what to do with other PathArguments"),
                };
                return (true, t);
            }
        }
    }
    (false, ty)
}
