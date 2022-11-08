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
    order: Option<u32>,
    width: Option<u32>,
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
        // println!("{:#?}", self);
        let struct_name = &self.name;
        let header_help_struct = self.generate_header_help_struct();
        let header_help_list_data = self.generate_header_help_data();

        // let fds = &self.fields;
        // let fields = &self.fields;
        let mut man_insert_stream = TokenStream::new();
        for fd in self.fields.iter() {
            if !fd.is_vec {
                if let Some(title) = &fd.opts.title {
                    let fd_name = &fd.name;
                    let title = format_title_with_dashed(title.clone(), &fd.name);
                    if fd.is_option {
                        man_insert_stream.extend(quote!(
                            map.insert(#title.to_string(),temp.#fd_name.clone().unwrap().to_string());
                        ));
                    } else {
                        man_insert_stream.extend(quote!(
                            map.insert(#title.to_string(),temp.#fd_name.clone().to_string());
                        ));
                    }
                    // let tt = title.split('-').next().unwrap().to_string();
                }
            }
        }
        let res = quote!(
            #header_help_struct

            impl #struct_name {
                fn write_excel(list: Vec<#struct_name>, path: &std::path::Path) {
                    use umya_spreadsheet as xl_tool;
                    use std::collections::HashMap;
                    let mut book = xl_tool::new_file_empty_worksheet();
                    let sheet = book.new_sheet("Sheet1").unwrap();
                    #header_help_list_data
                    // println!("{}",header_help_list.len());
                    for col in 1..=header_help_list.len() {
                        let cell = sheet.get_cell_by_column_and_row_mut(&(col as u32), &(1 as u32));
                        let title = &header_help_list[col - 1].title.clone().unwrap();
                        // let title = header[col - 1].split('-').next().unwrap().to_string();
                        cell.set_value_from_string(title);
                    }
                    for row in 2..=list.len()+1 {
                        let temp = &list[row - 2];
                        let mut map: HashMap<String,String> = HashMap::new();
                        #man_insert_stream
                        for col in 1..=header_help_list.len() {
                            let cell = sheet.get_cell_by_column_and_row_mut(&(col as u32), &(row as u32));
                            let title_key = &header_help_list[col - 1].title_key.clone().unwrap();
                            if let Some(value) = map.get(title_key) {
                                cell.set_value_from_string(value);
                            }else{
                                cell.set_value_from_string("".to_string());
                            }
                        }
                    }

                    // style width
                    for col in 1..=header_help_list.len() {
                        let column = sheet.get_column_dimension_by_number_mut(&(col as u32));
                        let width = header_help_list[col - 1].width.unwrap();
                        if width == 0 {
                            column.set_auto_width(true);
                        }else {
                            column.set_width(width as f64);
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
        let mut header_fd: Vec<_> = self
            .fields
            .iter()
            .filter(|fd| !fd.is_vec)
            .filter(|fd| fd.opts.title.is_some())
            .collect();

        let default_order = 99u32;
        header_fd.sort_by(|o1, o2| {
            let o1_order = o1.opts.order.unwrap_or(default_order);
            let o2_order = o2.opts.order.unwrap_or(default_order);
            o1_order.cmp(&o2_order)
        });

        let header: Vec<_> = header_fd
            .iter()
            .map(|fd| format_title_with_dashed(fd.opts.title.clone().unwrap(), &fd.name))
            .collect();
        header
    }

    fn generate_header_help_struct(&self) -> TokenStream {
        let header_help = Ident::new(&format!("{}HeaderHelp", self.name), self.name.span());
        let mut stream = TokenStream::new();
        stream.extend(quote!(
            #[derive(Debug)]
            struct #header_help{
                title: Option<String>,
                title_key: Option<String>,
                order: Option<u32>,
                width: Option<u32>,
            }
        ));
        stream
    }

    // const DEFAULT_ORDER: u32 = 99;
    fn generate_header_help_data(&self) -> TokenStream {
        let header_help_name = Ident::new(&format!("{}HeaderHelp", self.name), self.name.span());
        let mut header_struct_data_stream = TokenStream::new();
        header_struct_data_stream.extend(quote!(
            let mut header_help_list = Vec::new();
            // let mut header_help_map = HashMap::new();
        ));
        for fd in self.fields.iter() {
            if fd.is_vec {
                continue;
            }
            if let Some(title) = &fd.opts.title {
                let title_key = format_title_with_dashed(title.clone(), &fd.name);
                let order = fd.opts.order.unwrap_or(99);
                let width = fd.opts.width.unwrap_or(0);
                header_struct_data_stream.extend(quote!(
                    header_help_list.push(
                        #header_help_name {
                            title: Some(#title.to_string()),
                            title_key: Some(#title_key.to_string()),
                            order: Some(#order),
                            width: Some(#width),
                        }
                    );
                ));
            }
        }
        header_struct_data_stream.extend(quote!(
            header_help_list.sort_by(|o1,o2|{
                let o1_order = o1.order.unwrap();
                let o2_order = o2.order.unwrap();
                o1_order.cmp(&o2_order)
            });
            // println!("{:#?}", header_help_list);
        ));
        header_struct_data_stream
    }
}

fn format_title_with_dashed(title: String, fd_name: &Ident) -> String {
    format!("{}-{}", title, fd_name)
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
